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
vbwProfiler.vbwProcIn 247

    Dim oXDef As cXML
    Dim sXPath As String
    Dim hImageSmall As Long

vbwProfiler.vbwExecuteLine 4695
    On Error GoTo err
    '--
vbwProfiler.vbwExecuteLine 4696
    sXPath = oXML.GetXPath(oXNode)
    'uCmp.DefPath = App.Path & oXNode.Attributes.getNamedItem(XML_NODE_DEFPATH).nodeValue
vbwProfiler.vbwExecuteLine 4697
    uCmp.DefPath = App.Path & oXNode.selectSingleNode(XML_NODETYPE_STRING & "[@name='" & XML_NODE_DEFPATH & "']").nodeTypedValue
vbwProfiler.vbwExecuteLine 4698
    uCmp.GUID = oXNode.selectSingleNode(XML_NODETYPE_STRING & "[@name='" & XML_NODE_GUID & "']").nodeTypedValue
vbwProfiler.vbwExecuteLine 4699
    uCmp.Text = oXNode.selectSingleNode(XML_NODETYPE_STRING & "[@name='" & XML_NODE_NAME & "']").nodeTypedValue


    'Now we need to open the XML Defintion file
vbwProfiler.vbwExecuteLine 4700
    Set oXDef = New cXML

vbwProfiler.vbwExecuteLine 4701
    If oXDef.OpenFromFile(uCmp.DefPath, True) Then
vbwProfiler.vbwExecuteLine 4702
        Set oXNode = oXDef.GetRootNode
vbwProfiler.vbwExecuteLine 4703
        Set oXNode = oXNode.selectSingleNode(XML_NODE_OBJECT)

vbwProfiler.vbwExecuteLine 4704
        If uCmp.GUID = oXNode.selectSingleNode(XML_NODETYPE_STRING & "[@name='" & XML_NODE_GUID & "']").nodeTypedValue Then
vbwProfiler.vbwExecuteLine 4705
            uCmp.IconPath = App.Path & oXNode.selectSingleNode(XML_NODETYPE_STRING & "[@name='" & XML_NODE_IMAGE & "']").nodeTypedValue
vbwProfiler.vbwExecuteLine 4706
            uCmp.Classname = oXNode.selectSingleNode(XML_NODETYPE_STRING & "[@name='" & XML_NODE_CLASSNAME & "']").nodeTypedValue


            'note: We need to release the file handle it references or else
            '       if the file happens to be listed more than ONCE in the same
            '       vehicle, the next instance of XML reader wont be able to open it
vbwProfiler.vbwExecuteLine 4707
            Set oXDef = Nothing
vbwProfiler.vbwExecuteLine 4708
            ReadComponentXML = True

vbwProfiler.vbwExecuteLine 4709
            AddIconToImageList uCmp.IconPath
        Else
vbwProfiler.vbwExecuteLine 4710 'B
vbwProfiler.vbwExecuteLine 4711
            Debug.Print "modListViewHelper:ReadComponentXML() -- Error:  '" & uCmp.Text & ".cmp' GUID does not match it's .def GUID."
        End If
vbwProfiler.vbwExecuteLine 4712 'B
    End If
vbwProfiler.vbwExecuteLine 4713 'B
    '-
vbwProfiler.vbwProcOut 247
vbwProfiler.vbwExecuteLine 4714
    Exit Function
err:
vbwProfiler.vbwExecuteLine 4715
       Debug.Print "modListViewHelper:ReadComponentXML -- Error #" & err.Number & " " & err.Description
vbwProfiler.vbwExecuteLine 4716
       Resume Next
vbwProfiler.vbwProcOut 247
vbwProfiler.vbwExecuteLine 4717
End Function
                 
Sub LoadListView(ByVal sComponentType As String)
vbwProfiler.vbwProcIn 248
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
vbwProfiler.vbwExecuteLine 4718
    LockWindowUpdate frmDesigner.ListView1.hwnd
    ' Clear the listview
vbwProfiler.vbwExecuteLine 4719
    frmDesigner.ListView1.ListItems.Clear
vbwProfiler.vbwExecuteLine 4720
    frmDesigner.ListView1.SmallIcons = frmDesigner.ImageList1  '<-- add new icons to this list
vbwProfiler.vbwExecuteLine 4721
    frmDesigner.ListView1.Sorted = True
vbwProfiler.vbwExecuteLine 4722
    sPath = App.Path & "\components\" & sComponentType & "\" 'todo: need better use of constants for path
    ' start searching for components and add them.
    ' NOTE: It only displays the top level component even when a file contains nested components
vbwProfiler.vbwExecuteLine 4723
    sFileName = Dir(sPath & "*.cmp")


vbwProfiler.vbwExecuteLine 4724
    Set oXML = New cXML

'vbwLine 4725:    Do While sFileName <> ""
    Do While vbwProfiler.vbwExecuteLine(4725) Or sFileName <> ""
vbwProfiler.vbwExecuteLine 4726
        If oXML.OpenFromFile(sPath & sFileName, True) Then
vbwProfiler.vbwExecuteLine 4727
            Set oXRoot = oXML.GetRootNode()
vbwProfiler.vbwExecuteLine 4728
            Set oXNode = oXRoot.selectSingleNode(XML_NODE_OBJECT & "[@" & XML_ATTRIB_HANDLE & "='0_']")
            '--
vbwProfiler.vbwExecuteLine 4729
            If Not oXNode Is Nothing Then
vbwProfiler.vbwExecuteLine 4730
                Debug.Print "modListViewHelper.LoadListView() -- Found root object in CMP.XML file '" & sPath & sFileName & "'.  Attempting to load DEF.XML"
vbwProfiler.vbwExecuteLine 4731
                If ReadComponentXML(uComponent, oXNode, oXML) Then
                    ' all checks out, we can finally add it to the tree
vbwProfiler.vbwExecuteLine 4732
                    sIconKey = uComponent.IconPath

vbwProfiler.vbwExecuteLine 4733
                    sKey = sPath & sFileName
vbwProfiler.vbwExecuteLine 4734
                    uComponent.ComponentPath = sKey
vbwProfiler.vbwExecuteLine 4735
                    AddListViewItem sKey, uComponent.Text, sIconKey
                End If
vbwProfiler.vbwExecuteLine 4736 'B
            Else
vbwProfiler.vbwExecuteLine 4737 'B
vbwProfiler.vbwExecuteLine 4738
                Debug.Print "modListViewHelper.LoadListView() -- Could not find root object in XML file '" & sPath & sFileName & "'.  Most like caused by object handle not set to '0_'"
            End If
vbwProfiler.vbwExecuteLine 4739 'B
        End If
vbwProfiler.vbwExecuteLine 4740 'B
vbwProfiler.vbwExecuteLine 4741
        sFileName = Dir
vbwProfiler.vbwExecuteLine 4742
        uComponent.ComponentPath = sPath & sFileName
vbwProfiler.vbwExecuteLine 4743
    Loop

    'set the column widths
vbwProfiler.vbwExecuteLine 4744
    For i = 0 To frmDesigner.ListView1.ColumnHeaders.Count - 1
vbwProfiler.vbwExecuteLine 4745
        Call SendMessage(frmDesigner.ListView1.hwnd, LVM_SETCOLUMNWIDTH, i, 0)
vbwProfiler.vbwExecuteLine 4746
    Next

vbwProfiler.vbwExecuteLine 4747
    Set oXRoot = Nothing
vbwProfiler.vbwExecuteLine 4748
    Set oXNode = Nothing
vbwProfiler.vbwExecuteLine 4749
    LockWindowUpdate 0
vbwProfiler.vbwProcOut 248
vbwProfiler.vbwExecuteLine 4750
    Exit Sub

errorhandler:
vbwProfiler.vbwExecuteLine 4751
    LockWindowUpdate 0
vbwProfiler.vbwExecuteLine 4752
    Debug.Print "modListViewHelper:LoadListView -- Error " & error.Number & " " & error.Description

vbwProfiler.vbwProcOut 248
vbwProfiler.vbwExecuteLine 4753
End Sub

Public Function KeyFromLong(ByVal l As Long) As String
vbwProfiler.vbwProcIn 249
vbwProfiler.vbwExecuteLine 4754
    KeyFromLong = CStr(l) & "_"
vbwProfiler.vbwProcOut 249
vbwProfiler.vbwExecuteLine 4755
End Function

Private Function ImageListKeyExists(sKey As String) As Boolean
vbwProfiler.vbwProcIn 250
vbwProfiler.vbwExecuteLine 4756
    On Error GoTo errorhandler
    Dim s As String

vbwProfiler.vbwExecuteLine 4757
    s = frmDesigner.ImageList1.ListImages.Item(sKey).Tag
vbwProfiler.vbwExecuteLine 4758
    ImageListKeyExists = True
vbwProfiler.vbwProcOut 250
vbwProfiler.vbwExecuteLine 4759
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 4760
    If err.Number = 35601 Then
vbwProfiler.vbwExecuteLine 4761
        ImageListKeyExists = False
    Else
vbwProfiler.vbwExecuteLine 4762 'B
        ' we shouldnt be encountering any other error types here....
vbwProfiler.vbwExecuteLine 4763
        InfoPrint 1, "modListViewHelper:ImageListKeyExists -- Unexpected Error# " & err.Number & " " & err.Description
vbwProfiler.vbwExecuteLine 4764
        ImageListKeyExists = False
    End If
vbwProfiler.vbwExecuteLine 4765 'B
vbwProfiler.vbwProcOut 250
vbwProfiler.vbwExecuteLine 4766
End Function
Private Function StripExtensionFromFileName(ByVal sName As String) As String
vbwProfiler.vbwProcIn 251
    Dim lngLength As Long
    Dim lngExtensionPosition As Long

    Dim i As Long

vbwProfiler.vbwExecuteLine 4767
    lngLength = Len(sName)

vbwProfiler.vbwExecuteLine 4768
    For i = lngLength To 1 Step -1
vbwProfiler.vbwExecuteLine 4769
        If Mid(sName, i, 1) = "." Then
vbwProfiler.vbwExecuteLine 4770
             Exit For
        End If
vbwProfiler.vbwExecuteLine 4771 'B
vbwProfiler.vbwExecuteLine 4772
    Next

vbwProfiler.vbwExecuteLine 4773
    If (i = 1) And Mid(sName, 1, 1) <> "." Then
        ' there is no extension to strip off
vbwProfiler.vbwExecuteLine 4774
        StripExtensionFromFileName = sName
    Else
vbwProfiler.vbwExecuteLine 4775 'B
vbwProfiler.vbwExecuteLine 4776
        lngExtensionPosition = i - 1
vbwProfiler.vbwExecuteLine 4777
        StripExtensionFromFileName = Left(sName, lngExtensionPosition)
    End If
vbwProfiler.vbwExecuteLine 4778 'B
vbwProfiler.vbwProcOut 251
vbwProfiler.vbwExecuteLine 4779
End Function

Public Function CreateImageFromIconFile(sIconPath As String) As Long
vbwProfiler.vbwProcIn 252
    Dim lngRet As Long
    Dim hIcon As Long

vbwProfiler.vbwExecuteLine 4780
     hIcon = SHGetFileInfo(sIconPath, _
                0&, m_shinfo, Len(m_shinfo), _
                BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)

    'IMPORTANT: picIcon's width and height in pixels needs to be 16x16 else when Image_List draw is used
    'it will do some funky scaling and will screw the image up
    'frmDesigner.picIcon.ScaleMode = vbPixels
    'frmDesigner.picIcon.Height = 16
    'frmDesigner.picIcon.Width = 16
vbwProfiler.vbwExecuteLine 4781
    frmDesigner.picIcon.Picture = LoadPicture()
vbwProfiler.vbwExecuteLine 4782
    lngRet = ImageList_Draw(hIcon, m_shinfo.iIcon, frmDesigner.picIcon.hdc, 0, 0, ILD_TRANSPARENT)

vbwProfiler.vbwExecuteLine 4783
    frmDesigner.picIcon.Picture = frmDesigner.picIcon.Image

vbwProfiler.vbwExecuteLine 4784
    CreateImageFromIconFile = lngRet

vbwProfiler.vbwProcOut 252
vbwProfiler.vbwExecuteLine 4785
End Function

Public Function AddIconToImageList(ByRef sIconKey As String) As Long
vbwProfiler.vbwProcIn 253
    Dim imgX As ListImage
    Dim hImageSmall As Long
    Dim lRet As Long

vbwProfiler.vbwExecuteLine 4786
    If Not ImageListKeyExists(sIconKey) Then
vbwProfiler.vbwExecuteLine 4787
        hImageSmall = CreateImageFromIconFile(sIconKey)

vbwProfiler.vbwExecuteLine 4788
        Set imgX = frmDesigner.ImageList1.ListImages.Add(, sIconKey, frmDesigner.picIcon.Picture)
        'also add it to our treex
vbwProfiler.vbwExecuteLine 4789
        lRet = frmDesigner.treeVehicle.AddImage(sIconKey, frmDesigner.picIcon.Picture)
        'todo: do we need to ensure this got added to the tree's built in image list?
        ' do we need to check that its not already added to the tree's built in control?
        ' wouldn't it error if it were already in?  Only if "somehow" it was in the tree but magically not in the imagelist1
vbwProfiler.vbwExecuteLine 4790
        Debug.Print "modListViewHelper:AddIconToImageList -- " & lRet
    Else
vbwProfiler.vbwExecuteLine 4791 'B
vbwProfiler.vbwExecuteLine 4792
        Set imgX = frmDesigner.ImageList1.ListImages.Item(sIconKey)
    End If
vbwProfiler.vbwExecuteLine 4793 'B


vbwProfiler.vbwExecuteLine 4794
    Set imgX = Nothing
vbwProfiler.vbwExecuteLine 4795
    AddIconToImageList = frmDesigner.ImageList1.ListImages(sIconKey).index
vbwProfiler.vbwProcOut 253
vbwProfiler.vbwExecuteLine 4796
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 4797
    Set imgX = Nothing
vbwProfiler.vbwExecuteLine 4798
    AddIconToImageList = False
vbwProfiler.vbwProcOut 253
vbwProfiler.vbwExecuteLine 4799
End Function

Private Function AddListViewItem(ByRef sKey As String, sNodeName As String, sIconKey As String) As Long
vbwProfiler.vbwProcIn 254
    Dim itmX As ListItem

vbwProfiler.vbwExecuteLine 4800
    If ImageListKeyExists(sIconKey) Then
vbwProfiler.vbwExecuteLine 4801
        Set itmX = frmDesigner.ListView1.ListItems.Add(, sKey, sNodeName)
        'since the treeview and listview share the same imagelist, and uses the same filename=KEY convention
        'we dont have to worry about coordinating listview and treeview icons.
vbwProfiler.vbwExecuteLine 4802
        itmX.SmallIcon = frmDesigner.ImageList1.ListImages(sIconKey).Key
vbwProfiler.vbwExecuteLine 4803
        itmX.Tag = sIconKey
vbwProfiler.vbwExecuteLine 4804
        Set itmX = Nothing
vbwProfiler.vbwExecuteLine 4805
        AddListViewItem = True
    Else
vbwProfiler.vbwExecuteLine 4806 'B
vbwProfiler.vbwExecuteLine 4807
        AddListViewItem = False
    End If
vbwProfiler.vbwExecuteLine 4808 'B

vbwProfiler.vbwProcOut 254
vbwProfiler.vbwExecuteLine 4809
End Function

