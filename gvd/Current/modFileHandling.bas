Attribute VB_Name = "modFileHandling"
Option Explicit
Option Base 1

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'======================================================
' This is the UDT for the INI file
'======================================================
Private Type udtRegInfo
    RegName() As Byte
     RegNum() As Byte
     RegID As Long
End Type

Private RegInfo As udtRegInfo
'======================================================

'======================================================
' This is .VEH file  header information
'======================================================
Dim FormatSignature As Byte ' used with SIG_128 and SIG_129 to determine if this is a new file format and whether it has One or Two file headers (i.e. Header and Header2 UDT's)
Const SIG_128 = 10          ' only contains first header
Const SIG_129 = 11          ' contains first and second header
Const OFFSET1 = 14  'offset for start of second header
Const OFFSET2 = 1356 'offset for start of vehicle data
Dim m_lngOffset As Long

' GVD 1st header data
Private Type Header
    CRC32 As Long
    Major As Integer
    Minor As Integer
    Revision As Integer
    RegID As Integer
End Type

' 2nd header, contains user vehicle file info
Private Type Header2
   TL As Byte
   version As Single
   GUID As String * 39
   Category As String * 50
   subcategory As String * 50
   Name As String * 150         ' vehicle name and not FILENAME
   Class As String * 150
   author As String * 100
   email As String * 100
   url As String * 200
   jpgfilename As String * 255
   Description As String * 255 'max description that will be visible on site is only 255
End Type

'======================================================
' This is the UDT for the File Save/Load file
'======================================================

Private Type uComponent
    TreeInfo As String
    Properties() As String
End Type

Private Components() As uComponent
'========================================================


Const GVDLicenseFile = "GVD.lic"
Const GVDINIFile = "GVD.ini"

Const FLAG_NOZIP = 0 ' set to 0 for RELEASE builds

Private z As String

Public Sub ExportFile(sType As String)
    ' Code to export and view the file as either Text, HTML-classic gurps style or HTML-tables
    Dim Cancel As Boolean
    Dim sFileName As String
    Dim sExtension As String
    Dim sFilter As String
    Dim sTemp As String
    Dim oCDLG As clsCmdlg
    
    If sType = "Text" Or sType = "Text Slim" Then
        sExtension = ".txt"
        sFilter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    Else
        sExtension = ".html"
        sFilter = "HTML files (*.htm; *.html)|*.htm; *.html|All files (*.*)|*.*"
    End If
    On Error GoTo errorhandler
    Cancel = False
    With oCDLG
        ' todo: eventually use simpler code to check for existance?  dont really need to
        ' include the scripting runtime object if i stop using the filesystemobjects
        Dim oFile As FileSystemObject
        Set oFile = New FileSystemObject
        
        If sType = "Text" Then
            If oFile.FolderExists(Settings.TextExportPath) Then
                .InitialDir = Settings.TextExportPath
            Else
                .InitialDir = App.Path
            End If
        Else
            If oFile.FolderExists(Settings.HTMLExportPath) Then
                .InitialDir = Settings.HTMLExportPath
            Else
                .InitialDir = App.Path
            End If
        End If
        .DefaultFilename = ""
        '.DefaultExt = sExtension
        .Filter = sFilter
        .CancelError = True
        .MultiSelect = False
        '.flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    End With
    
    Cancel = oCDLG.ShowSave(frmDesigner.hwnd)
    If Not Cancel Then
        ' A fileName was selected. Add the code to save the file here
        sFileName = oCDLG.cFileName(0)
        
        'save the path
        If sType = "Text" Then
            Settings.TextExportPath = ExtractPathFromFile(sFileName)
        Else
            Settings.HTMLExportPath = ExtractPathFromFile(sFileName)
        End If
        DoEvents
    Else
        Exit Sub
    End If
    
    ' now generate the actual file
    Open sFileName For Output As #2
    sTemp = createGURPSText(sType)
    Print #2, sTemp
    Close #2
    
    Dim retval As Long
    Dim sProgramPath As String
    
    'make sure the path is set
    If sType = "Text" Or sType = "Text Slim" Then
        sProgramPath = Settings.TextViewerPath
    Else
        sProgramPath = Settings.HTMLBrowserPath
    End If
    
    If sProgramPath = "" Then
        MsgBox "No viewer specified."
        frmConfigure.Show vbModal, frmDesigner
        Set frmConfigure = Nothing
        
    Else 'attempt to launch the program
        retval = StartDoc(sProgramPath)
        If retval <= 32 Then        ' Error
            MsgBox "Web Page not Opened", vbExclamation, "URL Failed"
        End If
    End If
    Exit Sub

' check to see if the user Cancels out instead of electing to save a file
errorhandler:
    InfoPrint 1, "Error in ExportFile:  " + CStr(err.Number) + " " + err.Description
    Resume Next
End Sub

Function LoadRecords(ByVal FileName As String) As Boolean
' This function just loads the file data into an array of
' "components".  They are not actually added to the "Vehicle" at this point
' that is done in the RebuildComponentStructure function which is called at the end of this sub
' if the vehicle data was able to be read in properly

Dim sVersInfo() As String
Const Major = 0
Const Minor = 1
Const Revision = 2
Const REG_ID = 3

Dim iFreeFile As Long
Dim sLine As String


On Error GoTo errorhandler

' Destroy any old vehicle object and create new
Set Vehicle = Nothing
Set Vehicle = New Vehicles.clsVehicle
'todo: MUST use obtptr() of IComponent interface of Vehicle for Key in tree for Vehicle
' Clear the treeview of any nodes that might already exist
frmDesigner.treeVehicle.Nodes.Clear
z = m_oCurrentVeh.GetOverDrive

'//show the "loading" in the status bar
With frmDesigner
    .MousePointer = vbHourglass
    .ListView1.MousePointer = vbHourglass
End With

' set the status bar panels
frmDesigner.StatusBar1.Panels(1).Text = "Reading file data  0%"
frmDesigner.StatusBar1.Panels(1).Picture = frmDesigner.ImageList1.ListImages(11).Picture
    
' determine whether we are dealing with a old file format or new
If NewFileFormat(FileName) Then
    If LoadComponents_NewFormat(FileName) = False Then GoTo errorhandler
    Debug.Print "LoadRecords: GVD User Registration ID = " & gsRegID  'MPJ 07/04/2000
Else
    'open the file
    iFreeFile = FreeFile
    Open FileName For Input As #iFreeFile
    'get the first line which has the version info
    Line Input #iFreeFile, sLine
    Close #iFreeFile
    
    sLine = DecryptINI(DecryptINI$(sLine, z), z & Str(5982))  'the first line is double encrypted
    sVersInfo = Split(sLine, ",")
    gsMajor = sVersInfo(Major)
    gsMinor = sVersInfo(Minor)
    gsRevision = sVersInfo(Revision)
    
    If LoadComponents_OldFormat(FileName) = False Then GoTo errorhandler
End If

'//reset the status bar panels
frmDesigner.StatusBar1.Panels(1).Text = ""
frmDesigner.StatusBar1.Panels(1).Picture = LoadPicture()
    
'//Rebuild Vehicle Structure
If RebuildComponentStructure = False Then GoTo errorhandler

With frmDesigner
    .MousePointer = vbDefault
    .ListView1.MousePointer = vbDefault
    .treeVehicle.Nodes(BODY_KEY).Expanded = True
End With

'//load was successful
LoadRecords = True
Exit Function

errorhandler:
    With frmDesigner
        .MousePointer = vbDefault
        .ListView1.MousePointer = vbDefault
    End With
    
    MsgBox "Unable to load file.  Vehicle file is either invalid or corrupt."
    LoadRecords = False
    Close #iFreeFile
End Function

Function NewFileFormat(sFileName As String) As Boolean
    
    Dim iFree As Long
    Dim bSig As Byte
    Dim uHeader As Header
    Dim uHeader2 As Header2
    
    On Error GoTo errorhandler
    
    iFree = FreeFile
    Open sFileName For Binary As #iFree
    
    Get #iFree, , bSig
    
    ' returns True if bSig matches our Signature constant
    If (bSig <> SIG_128) And (bSig <> SIG_129) Then
        NewFileFormat = False
    Else
        'get the rest of the header
        Get #iFree, 2, uHeader
        With uHeader
            gsMajor = CStr(.Major)
            gsMinor = CStr(.Minor)
            gsRevision = CStr(.Revision)
        End With
            
        NewFileFormat = True
        
        ' set the appropriate offset for where our actual Vehicle data starts
        If bSig = SIG_129 Then
            m_lngOffset = OFFSET2
            
            ' get the second Header of the new file format so we can
            'obtain the GUID
            Get #iFree, OFFSET1, uHeader2
            p_sGUID = uHeader2.GUID
            
        ElseIf bSig = SIG_128 Then
            m_lngOffset = OFFSET1
        Else
            NewFileFormat = False
        End If
    End If

    Close #iFree
    Exit Function
    
errorhandler:
    NewFileFormat = False
    Close #iFree
End Function


Function LoadComponents_NewFormat(ByVal sFileName As String) As Boolean
     '//accepts a filename and loads in all the components
    Dim iFree As Long
    Dim sTemp As String
    Dim sArray() As String
    Dim b() As Byte
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim lngDataLen As Long
    Dim lngBytesRead As Long
    Dim oZip As cZlib
    Set oZip = New cZlib
    
    
    
    On Error GoTo errorhandler
    
    ReDim uRet(1)
    iFree = FreeFile
    
    
    Open sFileName For Binary As iFree
    
    'determine the length of the data
    lngDataLen = FileLen(sFileName) - m_lngOffset
    ReDim b(0 To lngDataLen - 1)
    
    'read it all in and decompress it
    Get #iFree, m_lngOffset, b
    
    If FLAG_NOZIP <> True Then
        oZip.UncompressB b
    End If
    
    ' convert to string and split this up into our seperate component lines
    sTemp = StrConv(b, vbUnicode)
    
    sArray = Split(sTemp, Chr(254))
        
  
   k = 0
    Do While k <= UBound(sArray)
               
        'fill the variant array that we will be passing into the FileLoader class
        If Left(sArray(k), 1) = "[" Then
            'now remove the leading and trailing characters
            i = i + 1
            j = 0
            ReDim Preserve Components(i)
            sArray(k) = Mid(sArray(k), 2, Len(sArray(k)) - 2)
            Components(i).TreeInfo = sArray(k)
            
        Else
            j = j + 1
            ReDim Preserve Components(i).Properties(j)
            Components(i).Properties(j) = sArray(k)
        End If
        k = k + 1
        frmDesigner.StatusBar1.Panels(1).Text = "Reading file data  " & CLng(k / UBound(sArray) * 100) & "%"
    Loop

    
    Close #iFree
    Set oZip = Nothing
    LoadComponents_NewFormat = True
    Debug.Print "LoadComponents_NewFormat: " & sTemp
    Exit Function
    
errorhandler:
    Debug.Print "LoadComponents_NewFormat: " & err.Description
    Set oZip = Nothing
    LoadComponents_NewFormat = False
End Function

Function LoadComponents_OldFormat(ByVal sFileName As String) As Boolean
    '//accepts a filename and loads in all the components
    Dim iFree As Long
    Dim sTemp As String
    Dim i As Long
    Dim j As Long
    Dim lngFileLen As Long
    Dim lngBytesRead As Long
    
    On Error GoTo errorhandler
    
    ReDim uRet(1)
    iFree = FreeFile
    
    lngFileLen = FileLen(sFileName)
    frmDesigner.StatusBar1.Panels(1).Text = "Reading file data  0%"
    frmDesigner.StatusBar1.Panels(1).Picture = frmDesigner.ImageList1.ListImages(11).Picture
    
    Open sFileName For Input As iFree
    
    'load the first line and skip it
    Line Input #iFree, sTemp
    
    
    Do While Not EOF(iFree)
        Line Input #iFree, sTemp
        lngBytesRead = lngBytesRead + Len(sTemp)
        sTemp = DecryptINI$(sTemp, z)
        'fill the variant array that we will be passing into the FileLoader class
        If Left(sTemp, 1) = "[" Then
            'now remove the leading and trailing characters
            i = i + 1
            j = 0
            ReDim Preserve Components(i)
            sTemp = Mid(sTemp, 2, Len(sTemp) - 2)
            Components(i).TreeInfo = sTemp
           
        Else
            j = j + 1
            ReDim Preserve Components(i).Properties(j)
            Components(i).Properties(j) = sTemp
        End If
        frmDesigner.StatusBar1.Panels(1).Text = "Reading file data  " & CLng(lngBytesRead / lngFileLen * 100) & "%"
    Loop

    
    Close #iFree

    LoadComponents_OldFormat = True
    Exit Function
    
errorhandler:
    
    LoadComponents_OldFormat = False
    
End Function

Function RebuildComponentStructure() As Boolean
    ' This is where the read in saved vehicle data is turned into the actual Vehicle heirarchy.
    ' It makes calls to Vehicle.AddObject for creating the correct objects based on the parsed datatypes
    Dim tobj As clsFileLoader
    Dim vc As Variant
    Dim A As Long
    Dim sKey As String
    Dim sParent As String
    Dim dType As Integer
    Dim memberID As String
    Dim propvalue As Variant
    Dim sDescription As String
    Dim icon1 As Integer
    Dim arrkey As Variant
    Dim i As Long
    Dim j As Long
    Dim iCount As Long
    Dim iPropCount As Long
    Dim lngUpper As Long
    
    On Error GoTo errorhandler
    '/show our progress meter
    lngUpper = UBound(Components)
    frmDesigner.StatusBar1.Panels(1).Text = "Building tree  0%"
    
    ' create an instance of the FileLoader class
    Set tobj = New clsFileLoader
    
    '//load up our tree and Component object
    For iCount = 1 To lngUpper
        vc = Split(Components(iCount).TreeInfo, "|")
        sDescription = vc(0)
        sKey = vc(1)
        sParent = vc(2)
        dType = vc(3)
        icon1 = vc(4)
        
         'create the vehicle component object
         If m_oCurrentVeh.addObject(dType, sKey, sParent, icon1, sDescription, True) Then
             'create the tree node (unless its like a weapon link, performance, etc)
             With frmDesigner.treeVehicle
                If sKey = BODY_KEY Then 'if its the body, then its the root node and doesnt have a parent
                    .Nodes.Add , , sKey, sDescription, icon1
                ElseIf (dType = PERFORMANCEPROFILE) Or (dType = WeaponLink) Then
                    'performance profiles and weapon links do NOT get added to the tree
                    'todo: They will have to now!
                Else
                    .Nodes.Add sParent, tvwChild, sKey, sDescription, icon1
                End If
            End With
        Else
            GoTo errorhandler
        End If
        
        For iPropCount = 1 To UBound(Components(iCount).Properties)
            vc = Split(Components(iCount).Properties(iPropCount), "|")
            'check for keychain.
            If vc(1) = "[" Then
                
                j = 1
                ReDim arrkey(UBound(vc) - 1)
                For i = 2 To UBound(vc)
                    arrkey(j) = vc(i)
                    j = j + 1
                Next
                propvalue = arrkey
                memberID = vc(0)
                tobj.LetProperties m_oCurrentVeh.Components(sKey), memberID, propvalue
            Else
                'fill the properties for this object
                Debug.Assert vc(0) <> "CombinedComponentVolume"
                memberID = vc(0)
                propvalue = vc(1)
                tobj.LetProperties m_oCurrentVeh.Components(sKey), memberID, propvalue
            End If
        Next
        
        frmDesigner.StatusBar1.Panels(1).Text = "Building tree  " & CLng(iCount / lngUpper * 100) & "%"
    Next
    
    RebuildComponentStructure = True
    'destroy the instance of the fileloader class
    Set tobj = Nothing
    frmDesigner.StatusBar1.Panels(1).Picture = LoadPicture()
    frmDesigner.StatusBar1.Panels(1).Text = ""
    Exit Function
    
errorhandler:
    Debug.Print "RebuildComponentStructure: " & err.Description
    frmDesigner.StatusBar1.Panels(1).Text = ""
    frmDesigner.StatusBar1.Panels(1).Picture = LoadPicture()
    RebuildComponentStructure = False
    
End Function
'//////////////////////////////////////////////////////
Function CreatePrintString(ByVal vc As Variant) As String()

Dim i, j As Long
Dim dType As Integer
Dim sKey As String
Dim sParent As String
Dim skeychain As String
Dim vType As Long
Dim svType As String
Dim sDescription As String
Dim icon1 As Integer
Dim icon2 As Integer
Dim retval() As String
Dim SIZE As Long

'find the key and datatype
For i = 1 To UBound(vc)
    If (vc(i, 0) = "Datatype") Or (vc(i, 0) = "datatype") Then
        dType = vc(0, i)
    ElseIf (vc(i, 0) = "Key") Or (vc(i, 0) = "key") Then
        sKey = vc(0, i)
        If sKey = BODY_KEY Then sParent = "0_"
    ElseIf (vc(i, 0) = "Parent") Or (vc(i, 0) = "parent") Then
        sParent = vc(0, i)
    ElseIf (vc(i, 0) = "customdescription") Or (vc(i, 0) = "Customdescription") Or (vc(i, 0) = "CustomDescription") Then
        sDescription = vc(0, i)
    ElseIf (vc(i, 0) = "Image") Or (vc(i, 0) = "image") Then
        icon1 = vc(0, i)
    ElseIf (vc(i, 0) = "SelectedImage") Or (vc(i, 0) = "selectedimage") Or (vc(i, 0) = "Selectedimage") Then
        icon2 = vc(0, i)
    End If
    If (dType <> vbEmpty) And (sKey <> "") And (sDescription <> "") And (icon1 <> vbEmpty) And (icon2 <> vbEmpty) And (sParent <> "") Then Exit For
Next

'store this line in the array
ReDim retval(1)
retval(1) = "[" + sDescription + "|" + sKey + "|" + sParent + "|" + Str(dType) + "|" + Str(icon1) + "|" + Str(icon2) + "]"

ReDim Preserve retval(UBound(vc) + 1)
SIZE = 1

For i = 1 To UBound(vc)
    SIZE = SIZE + 1
    vType = VarType(vc(0, i))
    svType = Str(vType)
    'check for keychains (vbarray + vbvariant)
    'reset the skeychain string
    skeychain = ""
    If vType = vbArray + vbVariant Then
        'found a keychain.  Place the "[" which inidcates a keychain
        skeychain = skeychain + "|" + "["
        'Seperate the individual keys into a string
        For j = 1 To UBound(vc(0, i))
           
            skeychain = skeychain + "|" + vc(0, i)(j)
        Next
        retval(SIZE) = vc(i, 0) + skeychain
    ElseIf vType = vbArray + vbString Then
        'found a keychain.  Place the "[" which inidcates a keychain
        skeychain = skeychain + "|" + "["
        'Seperate the individual keys into a string
        For j = 1 To UBound(vc(0, i))
           
            skeychain = skeychain + "|" + vc(0, i)(j)
        Next
        retval(SIZE) = vc(i, 0) + skeychain
    ElseIf vType <> vbString Then
        retval(SIZE) = vc(i, 0) + "|" + Str(vc(0, i))
    Else
        retval(SIZE) = vc(i, 0) + "|" + vc(0, i)
    End If
Next


CreatePrintString = retval
End Function

Sub CreateRecords()
    
    Dim tobj As clsFileLoader
    Dim vc As Variant
    Dim iIndex As Integer
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim keychainkeys As Variant
    Dim sKey As String
    
    On Error GoTo errorhandler
    
    ' create an instance of the FileLoader class
    Set tobj = New clsFileLoader
    z = m_oCurrentVeh.GetOverDrive
    
    'todo: Here instead of using GetFirstParent, we will start with the Body node and only
    ' itterate thru all children/sub children from this node.
    'NOTE: We do have to itterate thru the tree and not simply itterate thru the
    ' m_oCurrentVeh.components collection.  Itterating thru the components collection would result in
    'components being read and potentially the parent to which they must be added not being in the tree yet.
    'Proof: Lets say you add a Weapon to the Body and then a turret to the body.  Now lets say you move the weapon
    ' to the turret.  That puts the weapon ahead of the Turret in the collection so when reading in the weapon
    ' it would fail when trying to addobject to the parent turret which hasnt been installed yet.
    GetFirstParent 'Find a root node in the treeview
    'get the index of the root node that is at the top of the treeview
    iIndex = frmDesigner.treeVehicle.Nodes(p_nIndex).FirstSibling.index
    sKey = frmDesigner.treeVehicle.Nodes(iIndex).Key
    ' debug I really dont need a select case here since the first node is always the Body
    'sName = TypeName(m_oCurrentVeh.Components.item(sKey))
    vc = tobj.GetProperties(Vehicle.Components(sKey))
    i = 1
    ReDim Components(i)
    Components(1).Properties = CreatePrintString(vc)
    
    'If the Node has Children call the sub that writes the children
    If frmDesigner.treeVehicle.Nodes(iIndex).children > 0 Then
        WriteChild iIndex
    End If
    
    'Now save the Performance Profiles which are not
    'visually represented by the Tree
    keychainkeys = m_oCurrentVeh.Components(BODY_KEY).PerformanceProfileKeychain
    If (UBound(keychainkeys) >= 1) And keychainkeys(1) <> "" Then
        For k = 1 To UBound(keychainkeys)
            vc = tobj.GetProperties(m_oCurrentVeh.Components(keychainkeys(k)))
            i = UBound(Components) + 1
            ReDim Preserve Components(i)
            Components(i).Properties = CreatePrintString(vc)
        Next
    End If
    'Now save the Weapon Links which are not
    'visually represented by the Tree
    keychainkeys = m_oCurrentVeh.Components(BODY_KEY).WeaponLinkKeychain
    If (UBound(keychainkeys) >= 1) And keychainkeys(1) <> "" Then
        For k = 1 To UBound(keychainkeys)
            vc = tobj.GetProperties(m_oCurrentVeh.Components(keychainkeys(k)))
            i = UBound(Components) + 1
            ReDim Preserve Components(i)
            Components(i).Properties = CreatePrintString(vc)
        Next
    End If
    
    'destroy the instance of the fileloader class
    Set tobj = Nothing
    Exit Sub

errorhandler:

    
End Sub
Sub WriteRecord(ByVal FileName As String)
    'now get ready to print the properties for this object
'first delete the contents  of the file
Dim iFree As Long
Dim i, j As Long
Dim k As Long
Dim tempbyte() As Byte
Dim bFlag As Boolean
Dim lngUpper As Long
Dim s As String
Dim oZip As cZlib
Dim b() As Byte
Dim sJoin() As String
Dim uHeader As Header
Dim uHeader2 As Header2

iFree = FreeFile

'//reg check
tempbyte = ChopCheck
If (IsEmpty(tempbyte) = False) And (UBound(tempbyte) - LBound(tempbyte) = UBound(gsRegNum) - LBound(gsRegNum)) Then
    For i = 1 To UBound(gsRegNum)
        If tempbyte(i) = gsRegNum(i) Then
            bFlag = True
        Else
            bFlag = False
            GoTo reghandler
        End If
    Next
Else
    bFlag = False
    GoTo reghandler
End If

frmDesigner.StatusBar1.Panels(1).Text = "Reading vehicle data..."
CreateRecords

'delete the existing file
Open FileName For Random As #iFree
Close #iFree

frmDesigner.StatusBar1.Panels(1).Text = "Writing file  0%"
'open the file for writing
Open FileName For Binary As #iFree

Set oZip = New cZlib

lngUpper = UBound(Components)

ReDim sJoin(1)

For i = 1 To lngUpper
    'Print #iFree, EncryptINI$(Components(i).TreeInfo, z)
    's = s & Components(i).TreeInfo & Chr(254)
    If k > 0 Then
        ReDim Preserve sJoin(UBound(sJoin) + UBound(Components(i).Properties))
    Else
        ReDim sJoin(UBound(Components(i).Properties))
    End If
    For j = 1 To UBound(Components(i).Properties)
        k = k + 1
        sJoin(k) = Components(i).Properties(j)
        'Print #iFree, EncryptINI$(Components(i).Properties(j), z)

    Next
    
    frmDesigner.StatusBar1.Panels(1).Text = "Writing file  " & CLng(i / lngUpper * 100) & "%"
Next


'ReDim Preserve B(Len(B) - 1)

s = Join(sJoin, Chr(254))
'remove the last chr(254) from the end
's = Mid(s, 1, Len(s) - 1)
b = StrConv(s, vbFromUnicode)


If FLAG_NOZIP <> True Then
    oZip.CompressB b
End If

'create our file header
'first print the header version info
With uHeader
    .CRC32 = 100 'todo: need to calc the crc first
    .Major = App.Major
    .Minor = App.Minor
    .Revision = App.Revision
    .RegID = gsRegID
End With

' create our second file header
'todo:
'With m_oCurrentVeh.Components(BODY_KEY)
'   uHeader2.TL = .TL
'   uHeader2.version = m_oCurrentVeh.Description.version
'   'uHeader2.GUID = 'todo: Eh?  why commented out.  Cant remember but i do need to store that GUID right?
'   uHeader2.category = m_oCurrentVeh.Description.category
'   uHeader2.subcategory = m_oCurrentVeh.Description.subcategory
'   uHeader2.Class = m_oCurrentVeh.Description.Classname
'   uHeader2.name = m_oCurrentVeh.Description.NickName
'   uHeader2.author = m_oCurrentVeh.Description.author
'   uHeader2.email = m_oCurrentVeh.Description.email
'   uHeader2.url = m_oCurrentVeh.Description.url
'   uHeader2.jpgfilename = m_oCurrentVeh.Description.VehicleImageFileName
'   uHeader2.Description = m_oCurrentVeh.Description.VehicleDescription
'End With

' put the file header
Put #iFree, , SIG_129  'file version signature to help identify between all the legacy .veh formats
Put #iFree, 2, uHeader
Put #iFree, OFFSET1, uHeader2

' now put our actual data
Put #iFree, OFFSET2, b
Set oZip = Nothing
    
Close #iFree


' Writes the vehicle collection to a file specified by the user
reghandler:
    Close #iFree
    Exit Sub
End Sub

Private Sub WriteChild(ByVal iNodeIndex As Integer)
' Write the child nodes to the table. This sub uses recursion
' to loop through the child nodes. It receives the Index of
' the node that has the children
Dim tobj As clsFileLoader
Dim i As Long
Dim iTempIndex As Integer
Dim Temp() As String
Dim vc As Variant
Dim k As Long

' create an instance of the FileLoader class
Set tobj = New clsFileLoader

iTempIndex = frmDesigner.treeVehicle.Nodes(iNodeIndex).Child.FirstSibling.index
'Loop through all a Parents Child Nodes
i = frmDesigner.treeVehicle.Nodes(iNodeIndex).children
k = UBound(Components)

ReDim Preserve Components(k + i)

For i = 1 To frmDesigner.treeVehicle.Nodes(iNodeIndex).children
    vc = tobj.GetProperties(m_oCurrentVeh.Components(frmDesigner.treeVehicle.Nodes(iTempIndex).Key))
    k = k + 1
    Components(k).Properties = CreatePrintString(vc)
    
    'If the Node we are on has a child call the Sub again
    If frmDesigner.treeVehicle.Nodes(iTempIndex).children > 0 Then
       WriteChild (iTempIndex)
    End If
    
    'If we are not on the last child move to the next child Node
    If i <> frmDesigner.treeVehicle.Nodes(iNodeIndex).children Then
       iTempIndex = frmDesigner.treeVehicle.Nodes(iTempIndex).Next.index
    End If
Next

'destroy the instance of the fileloader class
Set tobj = Nothing

End Sub

Public Function PasteNode(sSourceKey() As String, sDestinationKey() As String, sKey() As String)
    '//to copy a node, you simply create a new object of the same type as the Source.
    Dim dType As Integer
    Dim sText As String
    Dim iImage1 As Integer
    Dim tobj As clsFileLoader
    Dim vc As Variant
    Dim i As Long
    Dim memberID As String
    Dim propvalue As Variant
    Dim iCount As Long
    Dim sOldKey As String
    Dim sNewKey As String
    Dim m As Long
    
    On Error GoTo errorhandler
    
    For iCount = 1 To UBound(sSourceKey)
                  
        With m_oCurrentVeh.Components(sSourceKey(iCount))
            dType = .Datatype
            sText = .CustomDescription
            iImage1 = .Image
            

        End With
        
        '//first add it to the tree
        frmDesigner.treeVehicle.Nodes.Add sDestinationKey(iCount), tvwChild, sKey(iCount), sText, iImage1
        '//now attempt to create the copied node and add it to the vehicle
        If m_oCurrentVeh.addObject(dType, sKey(iCount), sDestinationKey(iCount), iImage1, sText, True) = False Then
            frmDesigner.treeVehicle.Nodes.Remove (sKey(iCount))
            MsgBox "Error copying node. Copy/Paste operation aborted."
            Exit Function
        End If
        
        '//restore the key and parent key since it doesnt get properly applied
        '//in a paste operation
        m_oCurrentVeh.Components(sKey(iCount)).Parent = sDestinationKey(iCount)
        m_oCurrentVeh.Components(sKey(iCount)).Key = sKey(iCount)
        m_oCurrentVeh.Components(sKey(iCount)).Datatype = dType
        
        '//we need to run the location check ourselves since
        '//we are by passing it in the AddObject using the LoadedFlag = TRUE
        If m_oCurrentVeh.Components(sKey(iCount)).LocationCheck = False Then
            ' remove the object from the tree since the AddObject failed
            frmDesigner.treeVehicle.Nodes.Remove (sKey(iCount))
            m_oCurrentVeh.Components.Remove sKey(iCount)
            MsgBox "Invalid paste location for this component. Paste aborted."
            Exit Function
        End If
        
        Dim sLoc As String
        'get the location so we can restore it
        sLoc = m_oCurrentVeh.Components(sKey(iCount)).location
        
        '//now we need to restore all the values EXCEPT for the Parent and Key values
        '//crap and also power system values
         Set tobj = New clsFileLoader
         vc = tobj.GetProperties(m_oCurrentVeh.Components(sSourceKey(iCount)))
        
        For i = 1 To UBound(vc)
            memberID = vc(i, 0)
            propvalue = vc(0, i)
            If (memberID = "Location") Or (memberID = "Key") Or (memberID = "Parent") Then
            ElseIf (VarType(propvalue) = vbArray + vbVariant) Or (VarType(propvalue) = vbArray + vbString) Then
            Else
                tobj.LetProperties m_oCurrentVeh.Components(sKey(iCount)), memberID, propvalue
            End If
        Next
        
        'finally, we can add our keychain keys
         m_oCurrentVeh.keymanager.AddKeyChainKeys sKey(iCount)
    Next
    
    Exit Function
errorhandler:
    If err.Number = 35602 Then
        '//we need to change all the keys
        sNewKey = GetNextKey
        sOldKey = sKey(iCount)
        For m = iCount To UBound(sKey)
            If sKey(m) = sOldKey Then
                sKey(m) = sNewKey
            End If
        Next
        Resume
    End If
    
End Function


Function EncryptINI$(ByVal Strg$, ByVal Password$)
   Dim b$, s$, i As Integer, j As Integer
   Dim A1 As Integer, A2 As Integer, A3 As Integer, P$
   j = 1
   For i = 1 To Len(Password$)
     P$ = P$ & Asc(Mid$(Password$, i, 1))
   Next
    
   For i = 1 To Len(Strg$)
     A1 = Asc(Mid$(P$, j, 1))
     j = j + 1: If j > Len(P$) Then j = 1
     A2 = Asc(Mid$(Strg$, i, 1))
     A3 = A1 Xor A2
     b$ = Hex$(A3)
     If Len(b$) < 2 Then b$ = "0" + b$
     s$ = s$ + b$
   Next
   EncryptINI$ = s$
End Function

Function DecryptINI$(ByVal Strg$, ByVal Password$)
   Dim b$, s$, i As Integer, j As Integer
   Dim A1 As Integer, A2 As Integer, A3 As Integer, P$
   j = 1
   For i = 1 To Len(Password$)
     P$ = P$ & Asc(Mid$(Password$, i, 1))
   Next
   
   For i = 1 To Len(Strg$) Step 2
     A1 = Asc(Mid$(P$, j, 1))
     j = j + 1: If j > Len(P$) Then j = 1
     b$ = Mid$(Strg$, i, 2)
     A3 = Val("&H" + b$)
     A2 = A1 Xor A3
     s$ = s$ + Chr$(A2)
   Next
   DecryptINI$ = s$
End Function


Sub ReadINI()
    Dim oFile As FileSystemObject
    Dim oINI As cINI
    ' fill our settings with default values first in case some of the INI values are missing
    With Settings
        .InitialDir = GVDPath
        .DesktopX = Screen.Width
        .DesktopY = Screen.Height
        .windowstate = vbNormal
        .FormTop = 0
        .FormLeft = 0
        .FormHeight = 600
        .FormWidth = 800
        .Splitter1 = 200
        .Splitter2 = 340
        .HSplitter = 340
        .bUseSurfaceAreaTable = True
        .bUseDefaultTextViewer = True
        .bUseDefaultWebBrowser = True
        .AuthorName = ""
        .Copyright = ""
        .email = ""
        .url = ""
        .Header = ""
        .Footer = ""
        .PublishEmailAddress = "veh@makosoft.com" 'todo: 02/16/02 Is this ok to have hardcoded?
        .DecimalPlaces = 2
        .FormatString = "standard"
        '.bQuickStart = False '02/16/02 MPJ (OBSOLETE)
        '.bSoundOff = False   '02/16/02 MPJ (obsolete)
        .bAssociateExt = False
        .TextExportPath = GVDPath
        .HTMLExportPath = GVDPath
        .VehiclesOpenPath = GVDPath
        .VehiclesSavePath = GVDPath
    End With
    
    Set oFile = New FileSystemObject
    Set oINI = New cINI
    
    ' Make sure the INI file exists
    If oFile.FileExists(GVDPath & "\" & GVDINIFile) Then
    
        oINI.FileName = GVDPath & "\" & GVDINIFile
        
        'read in the settings and store them into our Settings UDT
        With Settings
        
            .TextExportPath = oINI.ReadString("Paths", "TextExportPath")
            .HTMLExportPath = oINI.ReadString("Paths", "HTMLExportPath")
            .VehiclesOpenPath = oINI.ReadString("Paths", "VehiclesOpenPath")
            .VehiclesSavePath = oINI.ReadString("Paths", "VehiclesSavePath")
        
            .Recent1 = oINI.ReadString("Recent", "Recent1")
            .Recent2 = oINI.ReadString("Recent", "Recent2")
            .Recent3 = oINI.ReadString("Recent", "Recent3")
            .Recent4 = oINI.ReadString("Recent", "Recent4")
            .Recent5 = oINI.ReadString("Recent", "Recent5")
            
            .HTMLBrowserPath = oINI.ReadString("Viewers", "HTMLBrowserPath")
            .TextViewerPath = oINI.ReadString("Viewers", "TextViewerPath")
     
            .DesktopX = oINI.ReadInteger("Display", "DesktopX")
            .DesktopY = oINI.ReadInteger("Display", "DesktopY")
            .windowstate = oINI.ReadInteger("Display", "State")
            .FormTop = oINI.ReadInteger("Display", "Top")
            .FormLeft = oINI.ReadInteger("Display", "Left")
            .FormWidth = oINI.ReadInteger("Display", "Width")
            .FormHeight = oINI.ReadInteger("Display", "Height")
            .Splitter1 = oINI.ReadInteger("Display", "Splitter1")
            If .Splitter1 <= 0 Then .Splitter1 = 4220
    
            .Splitter2 = oINI.ReadInteger("Display", "Splitter2")
            If .Splitter2 <= 0 Then .Splitter2 = 6720
            
            .HSplitter = oINI.ReadInteger("Display", "HSplitter")
            If .HSplitter <= 0 Then .HSplitter = 7000
            
            .bUseSurfaceAreaTable = oINI.ReadInteger("Config", "UseSurfaceAreaTable")
            .bUseDefaultTextViewer = oINI.ReadInteger("Config", "UseDefaultTextViewer")
            .bUseDefaultWebBrowser = oINI.ReadInteger("Config", "UseDefaultWebBrowser")
            .PublishEmailAddress = oINI.ReadString("Config", "PublishEmail")
            .DecimalPlaces = oINI.ReadInteger("Config", "DecimalPlaces")
            .FormatString = oINI.ReadString("Config", "FormatCode")
            '.bQuickStart = oINI.ReadInteger("Config", "QuickStart") '02/16/02 MPJ (obsolete)
            '.bSoundOff = oINI.ReadInteger("Config", "DisableSound") '02/16/02 MPJ (obsolete)
            .bAssociateExt = oINI.ReadInteger("Config", "AssociateExt")
                
            .AuthorName = oINI.ReadString("Author", "Name")
            .email = oINI.ReadString("Author", "Email")
            .url = oINI.ReadString("Author", "URL")
            .Copyright = oINI.ReadString("Author", "Copyright")
            .Header = oINI.ReadString("Author", "Header")
            .Footer = oINI.ReadString("Author", "Footer")
        End With
    End If


End Sub

Public Sub ReadLicenseFile()
    Dim oFile As FileSystemObject
    Dim iFree As Long
    
    On Error GoTo errorhandler
    
    Set oFile = New FileSystemObject
    
    iFree = FreeFile
    
    If oFile.FileExists(GVDPath & "\" & GVDLicenseFile) Then   ' debug  this line needs to change because its saving the INI in the wrong place
        'retreive the data from the file and store it into the Settings udt
        Open GVDPath + "\" + GVDLicenseFile For Binary As #iFree
        Get #iFree, 1, RegInfo
        Close #iFree
    
        With RegInfo
            gsRegName = .RegName
            gsRegID = .RegID
            gsRegNum = .RegNum
        End With
    End If
    
    If IsEmpty(gsRegName) Then
        ReDim gsRegName(1)
    End If
    If IsEmpty(gsRegNum) Then
        ReDim gsRegNum(1)
    End If
    Exit Sub
    
errorhandler:
    ' must make sure our byte arrays are filled
    ReDim gsRegName(1)
    ReDim gsRegNum(1)
    
    'first close the file
    Close #iFree
    ' now delete it
    If oFile.FileExists(GVDPath & "\" & GVDLicenseFile) Then
    
        oFile.DeleteFile (GVDPath & "\" & GVDLicenseFile)
    End If
    
    DoEvents
    
End Sub

Public Sub WriteLicenseFile()
    
    Dim iFree As Long

    iFree = FreeFile

    'first delete the file if it already exists
    Open GVDPath + "\" + GVDLicenseFile For Random As #iFree
    Close #iFree
    
    iFree = FreeFile
    
    ' open the License for binary write
    Open GVDPath + "\" + GVDLicenseFile For Binary As #iFree
    
    ' update the relevant settings before we save it
    With RegInfo
        .RegID = gsRegID
        .RegName = gsRegName
        .RegNum = gsRegNum
    End With
     
    ' save the Settings data and close the file
    Put #iFree, , RegInfo
    Close #iFree

End Sub

Sub WriteINI()
    On Error Resume Next
    Dim oINI As cINI
    Set oINI = New cINI
    
    oINI.FileName = GVDPath & "\" & GVDINIFile
    
    With Settings
        Call oINI.WriteInteger("Display", "DesktopX", Screen.Width)
        Call oINI.WriteInteger("Display", "DesktopY", Screen.Height)
        Call oINI.WriteInteger("Display", "State", frmDesigner.windowstate)
        
        'JAW 2000.05.22
        'Splitter positions were not being saved when window was maximized.
        'If (frmDesigner.windowstate <> vbMaximized) And (frmDesigner.windowstate <> vbMinimized) Then
        If (frmDesigner.windowstate <> vbMinimized) Then
            Call oINI.WriteInteger("Display", "Top", Settings.FormTop)
            Call oINI.WriteInteger("Display", "Left", Settings.FormLeft)
            Call oINI.WriteInteger("Display", "Width", Settings.FormWidth)
            Call oINI.WriteInteger("Display", "Height", Settings.FormHeight)
            Call oINI.WriteInteger("Display", "Splitter1", Settings.Splitter1)
           ' Call oINI.WriteInteger("Display", "Splitter2", frmDesigner.ListView1.Left + frmDesigner.ListView1.Width)
            Call oINI.WriteInteger("Display", "HSplitter", Settings.HSplitter)
        End If
        
        Call oINI.WriteInteger("Config", "UseSurfaceAreaTable", Settings.bUseSurfaceAreaTable)
        Call oINI.WriteInteger("Config", "UseDefaultTextViewer", Settings.bUseDefaultTextViewer)
        Call oINI.WriteInteger("Config", "UseDefaultWebBrowser", Settings.bUseDefaultWebBrowser)
        Call oINI.WriteString("Config", "PublishEmail", .PublishEmailAddress)
        Call oINI.WriteInteger("Config", "DecimalPlaces", .DecimalPlaces)
        Call oINI.WriteString("Config", "FormatCode", "standard")
       ' Call oINI.WriteInteger("Config", "QuickStart", Settings.bQuickStart) '02/16/02 MPJ (obsolete)
       ' Call oINI.WriteInteger("Config", "DisableSound", Settings.bSoundOff) '02/16/02 MPJ (obsolete)
        Call oINI.WriteInteger("Config", "AssociateExt", Settings.bAssociateExt)
            
        Call oINI.WriteString("Author", "Name", Settings.AuthorName)
        Call oINI.WriteString("Author", "Email", Settings.email)
        Call oINI.WriteString("Author", "URL", Settings.url)
        Call oINI.WriteString("Author", "Copyright", Settings.Copyright)
        Call oINI.WriteString("Author", "Header", Settings.Header)
        Call oINI.WriteString("Author", "Footer", Settings.Footer)
        
        Call oINI.WriteString("Paths", "App", GVDPath)
        Call oINI.WriteString("Paths", "TextExportPath", .TextExportPath)
        Call oINI.WriteString("Paths", "HTMLExportPath", .HTMLExportPath)
        Call oINI.WriteString("Paths", "VehiclesOpenPath", .VehiclesOpenPath)
        Call oINI.WriteString("Paths", "VehiclesSavePath", .VehiclesSavePath)
        
        Call oINI.WriteString("Viewers", "HTMLBrowserPath", .HTMLBrowserPath)
        Call oINI.WriteString("Viewers", "TextViewerPath", .TextViewerPath)
     
        Call oINI.WriteString("Recent", "Recent1", frmDesigner.mnuRecent(1).Caption)
        Call oINI.WriteString("Recent", "Recent2", frmDesigner.mnuRecent(2).Caption)
        Call oINI.WriteString("Recent", "Recent3", frmDesigner.mnuRecent(3).Caption)
        Call oINI.WriteString("Recent", "Recent4", frmDesigner.mnuRecent(4).Caption)
        Call oINI.WriteString("Recent", "Recent5", frmDesigner.mnuRecent(5).Caption)
    End With
End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''OBSOLETE MPJ Oct.6.2002
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''Sub SaveComponent(ByVal sKey As String, ByVal FileName As String)
''''''''takes the Key and FileName of a weapon component and saves it
'''''''' Writes the vehicle collection to a file specified by the user
'''''''
'''''''Dim tobj As clsFileLoader
'''''''Dim vc As Variant
'''''''Dim Temp() As String
'''''''Dim i As Long
'''''''Dim iFree As Long
'''''''
'''''''On Error GoTo errorhandler
'''''''
'''''''' create an instance of the FileLoader class
'''''''Set tobj = New clsFileLoader
'''''''z = m_oCurrentVeh.GetOverDrive
'''''''
''''''''now get ready to print the properties for this object
''''''''first delete the contents  of the file
'''''''iFree = FreeFile
'''''''Open FileName For Random As #iFree
'''''''Close #iFree
'''''''
''''''''open the file for writing
'''''''Open FileName For Output As #iFree
'''''''
''''''''print the first line
'''''''With m_oCurrentVeh.Components(sKey)
'''''''    Print #iFree, .Datatype, .Image, .SelectedImage, .CustomDescription
'''''''End With
''''''''output it all
'''''''vc = tobj.GetProperties(m_oCurrentVeh.Components(sKey))
'''''''Temp = CreatePrintString(vc)
'''''''
'''''''
'''''''For i = 1 To UBound(Temp)
'''''''
'''''''    Print #iFree, EncryptINI$(Temp(i), z)
'''''''Next
'''''''
''''''''close the file
'''''''Close #iFree
'''''''
''''''''destroy the instance of the fileloader class
'''''''Set tobj = Nothing
'''''''Exit Sub
'''''''
'''''''errorhandler:
'''''''
'''''''    Close #iFree
'''''''
'''''''End Sub
'''''''
'''''''Function RestoreSavedItem(ByVal sFileName As String, ByVal sKey As String, ByVal sParent As String, ByRef sLocation As String)
''''''''loads a saved component with a given filename and
''''''''restores its properties
'''''''    Dim tobj As clsFileLoader
'''''''    Dim vc As Variant
'''''''    Dim memberID As String
'''''''    Dim propvalue As Variant
'''''''    Dim arrkey As Variant
'''''''    Dim i As Long
'''''''    Dim j As Long
'''''''    Dim iCount As Long
'''''''    Dim iPropCount As Long
'''''''    Dim lngUpper As Long
'''''''    Dim iFree As Long
'''''''    Dim sTemp As String
'''''''
'''''''    On Error GoTo errorhandler
'''''''    z = m_oCurrentVeh.GetOverDrive
'''''''    ReDim uRet(1)
'''''''    iFree = FreeFile
'''''''
'''''''    Open sFileName For Input As iFree
'''''''
'''''''     '//load the file data
'''''''    Line Input #iFree, sTemp '//skip the first line
'''''''    Do While Not EOF(iFree)
'''''''        Line Input #iFree, sTemp
'''''''        sTemp = DecryptINI$(sTemp, z)
'''''''        'fill the variant array that we will be passing into the FileLoader class
'''''''        If Left(sTemp, 1) = "[" Then
'''''''            'now remove the leading and trailing characters
'''''''            i = i + 1
'''''''            j = 0
'''''''            ReDim Preserve Components(i)
'''''''            sTemp = Mid(sTemp, 2, Len(sTemp) - 2)
'''''''            Components(i).TreeInfo = sTemp
'''''''
'''''''        Else
'''''''            j = j + 1
'''''''            ReDim Preserve Components(i).Properties(j)
'''''''            Components(i).Properties(j) = sTemp
'''''''        End If
'''''''    Loop
'''''''    Close #iFree
'''''''
'''''''   '//restore the values
'''''''    ' create an instance of the FileLoader class
'''''''    Set tobj = New clsFileLoader
'''''''
'''''''    For iPropCount = 1 To UBound(Components(1).Properties)
'''''''        vc = Split(Components(1).Properties(iPropCount), "|")
'''''''        'do NOT restore keychains
'''''''        If vc(1) = "[" Then
'''''''
'''''''        Else
'''''''            'fill the properties for this object
'''''''            Debug.Assert vc(0) <> "CombinedComponentVolume"
'''''''            memberID = vc(0)
'''''''            propvalue = vc(1)
'''''''            '//we dont want to restore non relevant values
'''''''            Select Case memberID
'''''''                Case "Parent", "Key", "Location", "ParentDatatype", "PrintOutput"  '<-- Property Exclusion
'''''''                Case Else
'''''''                    tobj.LetProperties m_oCurrentVeh.Components(sKey), memberID, propvalue
'''''''            End Select
'''''''        End If
'''''''    Next
'''''''
'''''''    ' set the actual key and parent key and location ' Though we exclude these out from above its probably not even necessary
'''''''     'set the new key!
'''''''    With m_oCurrentVeh.Components(sKey)
'''''''        .Key = sKey
'''''''        .Parent = sParent
'''''''        .location = sLocation '//restore the user specified location since it got overwritten after restoring attributes+stats from disk
'''''''    End With
'''''''
'''''''    ' 05/18/02 MPJ - Since AddObject now handles this, no need to AddKeyChainKeys are it results in TWO sets
'''''''    ' of keys being added.
'''''''    'finally, we can add our relevant keychain keys
'''''''    'm_oCurrentVeh.keymanager.AddKeyChainKeys sKey
'''''''
'''''''Exit Function
'''''''errorhandler:
'''''''
'''''''
'''''''End Function

