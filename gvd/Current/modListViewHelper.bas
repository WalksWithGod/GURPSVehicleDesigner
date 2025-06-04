Attribute VB_Name = "modListViewHelper"
Option Explicit

'constants and declarations for the listview column resizing
Private Const LVM_First As Long = &H1000
Private Const LVSCW_AUTOSIZE = -1
Private Const LVM_SETCOLUMNWIDTH As Long = LVM_First + 30
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
    
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000  'system icon index
Private Const SHGFI_LARGEICON = &H0        'large icon
Private Const SHGFI_SMALLICON = &H1        'small icon
Private m_shinfo As SHFILEINFO

Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or _
             SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or _
             SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Const MAX_PATH = 255
Private Const ILD_TRANSPARENT = &H1        'display transparent

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" _
   (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long
    
Private Type SHFILEINFO
   hIcon          As Long
   iIcon          As Long
   dwAttributes   As Long
   szDisplayName  As String * MAX_PATH
   szTypeName     As String * 80
End Type


Private Declare Function ImageList_Draw Lib "comctl32.dll" _
   (ByVal himl&, _
    ByVal i&, _
    ByVal hDCDest&, _
    ByVal x&, _
    ByVal y&, _
    ByVal flags&) As Long

Public Function ReadComponentXML(uCmp As udtComponent, oXNode As MSXML2.IXMLDOMNode, oXML As cXML) As Boolean
    
    Dim oXDef As cXML
    Dim sXPath As String
    Dim hImageSmall As Long
    
    On Error GoTo err
    '--
    sXPath = oXML.GetXPath(oXNode)
    'uCmp.DefPath = App.Path & oXNode.Attributes.getNamedItem(XML_NODE_DEFPATH).nodeValue
    uCmp.DefPath = App.Path & oXNode.selectSingleNode(XML_NODETYPE_STRING & "[@name='" & XML_NODE_DEFPATH & "']").nodeTypedValue
    uCmp.GUID = oXNode.selectSingleNode(XML_NODETYPE_STRING & "[@name='" & XML_NODE_GUID & "']").nodeTypedValue
    uCmp.Text = oXNode.selectSingleNode(XML_NODETYPE_STRING & "[@name='" & XML_NODE_NAME & "']").nodeTypedValue
    
    
    'Now we need to open the XML Defintion file
    Set oXDef = New cXML
    
    If oXDef.OpenFromFile(uCmp.DefPath, True) Then
        Set oXNode = oXDef.GetRootNode
        Set oXNode = oXNode.selectSingleNode(XML_NODE_OBJECT)
    
        If uCmp.GUID = oXNode.selectSingleNode(XML_NODETYPE_STRING & "[@name='" & XML_NODE_GUID & "']").nodeTypedValue Then
            uCmp.IconPath = App.Path & oXNode.selectSingleNode(XML_NODETYPE_STRING & "[@name='" & XML_NODE_IMAGE & "']").nodeTypedValue
            uCmp.Classname = oXNode.selectSingleNode(XML_NODETYPE_STRING & "[@name='" & XML_NODE_CLASSNAME & "']").nodeTypedValue
            
                        
            'note: We need to release the file handle it references or else
            '       if the file happens to be listed more than ONCE in the same
            '       vehicle, the next instance of XML reader wont be able to open it
            Set oXDef = Nothing
            ReadComponentXML = True
            
            AddIconToImageList uCmp.IconPath
        Else
            Debug.Print "modListViewHelper:ReadComponentXML() -- Error:  '" & uCmp.Text & ".cmp' GUID does not match it's .def GUID."
        End If
    End If
    '-
    Exit Function
err:
       Debug.Print "modListViewHelper:ReadComponentXML -- Error #" & err.Number & " " & err.Description
       Resume Next
End Function
                 
Sub LoadListView(ByVal sComponentType As String)
    Dim i As Long
    Dim sFileName As String
    Dim sPath As String
    Dim sKey As String
    Dim sIconKey As String
    Dim uComponent As udtComponent
    '--------
    Dim oXML As cXML
    Dim oXRoot As MSXML2.IXMLDOMNode
    Dim oXNode As MSXML2.IXMLDOMNode
    
    ' prevent listview updates til we're done updating
    LockWindowUpdate frmDesigner.ListView1.hwnd
    ' Clear the listview
    frmDesigner.ListView1.ListItems.Clear
    frmDesigner.ListView1.SmallIcons = frmDesigner.ImageList1  '<-- add new icons to this list
    frmDesigner.ListView1.Sorted = True
    sPath = App.Path & "\components\" & sComponentType & "\" 'todo: need better use of constants for path
    ' start searching for components and add them.
    ' NOTE: It only displays the top level component even when a file contains nested components
    sFileName = Dir(sPath & "*.cmp")
    
    
    Set oXML = New cXML
    
    Do While sFileName <> ""
        If oXML.OpenFromFile(sPath & sFileName, True) Then
            Set oXRoot = oXML.GetRootNode()
            Set oXNode = oXRoot.selectSingleNode(XML_NODE_OBJECT & "[@" & XML_ATTRIB_HANDLE & "='0_']")
            '--
            If Not oXNode Is Nothing Then
                Debug.Print "modListViewHelper.LoadListView() -- Found root object in CMP.XML file '" & sPath & sFileName & "'.  Attempting to load DEF.XML"
                If ReadComponentXML(uComponent, oXNode, oXML) Then
                    ' all checks out, we can finally add it to the tree
                    sIconKey = uComponent.IconPath
                    
                    sKey = sPath & sFileName
                    uComponent.ComponentPath = sKey
                    AddListViewItem sKey, uComponent.Text, sIconKey
                End If
            Else
                Debug.Print "modListViewHelper.LoadListView() -- Could not find root object in XML file '" & sPath & sFileName & "'.  Most like caused by object handle not set to '0_'"
            End If
        End If
        sFileName = Dir
        uComponent.ComponentPath = sPath & sFileName
    Loop
            
    'set the column widths
    For i = 0 To frmDesigner.ListView1.ColumnHeaders.Count - 1
        Call SendMessage(frmDesigner.ListView1.hwnd, LVM_SETCOLUMNWIDTH, i, 0)
    Next
    
    Set oXRoot = Nothing
    Set oXNode = Nothing
    LockWindowUpdate 0
    Exit Sub

errorhandler:
    LockWindowUpdate 0
    Debug.Print "modListViewHelper:LoadListView -- Error " & error.Number & " " & error.Description
    
End Sub

Public Function KeyFromLong(ByVal l As Long) As String
    KeyFromLong = CStr(l) & "_"
End Function

Private Function ImageListKeyExists(sKey As String) As Boolean
    On Error GoTo errorhandler
    Dim s As String
    
    s = frmDesigner.ImageList1.ListImages.Item(sKey).Tag
    ImageListKeyExists = True
    Exit Function
errorhandler:
    If err.Number = 35601 Then
        ImageListKeyExists = False
    Else
        ' we shouldnt be encountering any other error types here....
        InfoPrint 1, "modListViewHelper:ImageListKeyExists -- Unexpected Error# " & err.Number & " " & err.Description
        ImageListKeyExists = False
    End If
End Function
Private Function StripExtensionFromFileName(ByVal sName As String) As String
    Dim lngLength As Long
    Dim lngExtensionPosition As Long
    
    Dim i As Long

    lngLength = Len(sName)
    
    For i = lngLength To 1 Step -1
        If Mid(sName, i, 1) = "." Then Exit For
    Next
        
    If (i = 1) And Mid(sName, 1, 1) <> "." Then
        ' there is no extension to strip off
        StripExtensionFromFileName = sName
    Else
        lngExtensionPosition = i - 1
        StripExtensionFromFileName = Left(sName, lngExtensionPosition)
    End If
End Function

Public Function CreateImageFromIconFile(sIconPath As String) As Long
    Dim lngRet As Long
    Dim hIcon As Long
    
     hIcon = SHGetFileInfo(sIconPath, _
                0&, m_shinfo, Len(m_shinfo), _
                BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
                
    'IMPORTANT: picIcon's width and height in pixels needs to be 16x16 else when Image_List draw is used
    'it will do some funky scaling and will screw the image up
    'frmDesigner.picIcon.ScaleMode = vbPixels
    'frmDesigner.picIcon.Height = 16
    'frmDesigner.picIcon.Width = 16
    frmDesigner.picIcon.Picture = LoadPicture()
    lngRet = ImageList_Draw(hIcon, m_shinfo.iIcon, frmDesigner.picIcon.hdc, 0, 0, ILD_TRANSPARENT)
    
    frmDesigner.picIcon.Picture = frmDesigner.picIcon.Image
    
    CreateImageFromIconFile = lngRet
      
End Function

Public Function AddIconToImageList(ByRef sIconKey As String) As Long
    Dim imgX As ListImage
    Dim hImageSmall As Long
    Dim lRet As Long
    
    If Not ImageListKeyExists(sIconKey) Then
        hImageSmall = CreateImageFromIconFile(sIconKey)
        
        Set imgX = frmDesigner.ImageList1.ListImages.Add(, sIconKey, frmDesigner.picIcon.Picture)
        'also add it to our treex
        lRet = frmDesigner.treeVehicle.AddImage(sIconKey, frmDesigner.picIcon.Picture)
        'todo: do we need to ensure this got added to the tree's built in image list?
        ' do we need to check that its not already added to the tree's built in control?
        ' wouldn't it error if it were already in?  Only if "somehow" it was in the tree but magically not in the imagelist1
        Debug.Print "modListViewHelper:AddIconToImageList -- " & lRet
    Else
        Set imgX = frmDesigner.ImageList1.ListImages.Item(sIconKey)
    End If
    
    
    Set imgX = Nothing
    AddIconToImageList = frmDesigner.ImageList1.ListImages(sIconKey).index
    Exit Function
errorhandler:
    Set imgX = Nothing
    AddIconToImageList = False
End Function

Private Function AddListViewItem(ByRef sKey As String, sNodeName As String, sIconKey As String) As Long
    Dim itmX As ListItem
    
    If ImageListKeyExists(sIconKey) Then
        Set itmX = frmDesigner.ListView1.ListItems.Add(, sKey, sNodeName)
        'since the treeview and listview share the same imagelist, and uses the same filename=KEY convention
        'we dont have to worry about coordinating listview and treeview icons.
        itmX.SmallIcon = frmDesigner.ImageList1.ListImages(sIconKey).Key
        itmX.Tag = sIconKey
        Set itmX = Nothing
        AddListViewItem = True
    Else
        AddListViewItem = False
    End If
    
End Function
