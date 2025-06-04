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
vbwProfiler.vbwProcIn 12
    Dim Cancel As Boolean
    Dim sFileName As String
    Dim sExtension As String
    Dim sFilter As String
    Dim sTemp As String
    Dim oCDLG As clsCmdlg

vbwProfiler.vbwExecuteLine 267
    If sType = "Text" Or sType = "Text Slim" Then
vbwProfiler.vbwExecuteLine 268
        sExtension = ".txt"
vbwProfiler.vbwExecuteLine 269
        sFilter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    Else
vbwProfiler.vbwExecuteLine 270 'B
vbwProfiler.vbwExecuteLine 271
        sExtension = ".html"
vbwProfiler.vbwExecuteLine 272
        sFilter = "HTML files (*.htm; *.html)|*.htm; *.html|All files (*.*)|*.*"
    End If
vbwProfiler.vbwExecuteLine 273 'B
vbwProfiler.vbwExecuteLine 274
    On Error GoTo errorhandler
vbwProfiler.vbwExecuteLine 275
    Cancel = False
vbwProfiler.vbwExecuteLine 276
    With oCDLG
        ' todo: eventually use simpler code to check for existance?  dont really need to
        ' include the scripting runtime object if i stop using the filesystemobjects
        Dim oFile As FileSystemObject
vbwProfiler.vbwExecuteLine 277
        Set oFile = New FileSystemObject

vbwProfiler.vbwExecuteLine 278
        If sType = "Text" Then
vbwProfiler.vbwExecuteLine 279
            If oFile.FolderExists(Settings.TextExportPath) Then
vbwProfiler.vbwExecuteLine 280
                .InitialDir = Settings.TextExportPath
            Else
vbwProfiler.vbwExecuteLine 281 'B
vbwProfiler.vbwExecuteLine 282
                .InitialDir = App.Path
            End If
vbwProfiler.vbwExecuteLine 283 'B
        Else
vbwProfiler.vbwExecuteLine 284 'B
vbwProfiler.vbwExecuteLine 285
            If oFile.FolderExists(Settings.HTMLExportPath) Then
vbwProfiler.vbwExecuteLine 286
                .InitialDir = Settings.HTMLExportPath
            Else
vbwProfiler.vbwExecuteLine 287 'B
vbwProfiler.vbwExecuteLine 288
                .InitialDir = App.Path
            End If
vbwProfiler.vbwExecuteLine 289 'B
        End If
vbwProfiler.vbwExecuteLine 290 'B
vbwProfiler.vbwExecuteLine 291
        .DefaultFilename = ""
        '.DefaultExt = sExtension
vbwProfiler.vbwExecuteLine 292
        .Filter = sFilter
vbwProfiler.vbwExecuteLine 293
        .CancelError = True
vbwProfiler.vbwExecuteLine 294
        .MultiSelect = False
        '.flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
vbwProfiler.vbwExecuteLine 295
    End With

vbwProfiler.vbwExecuteLine 296
    Cancel = oCDLG.ShowSave(frmDesigner.hwnd)
vbwProfiler.vbwExecuteLine 297
    If Not Cancel Then
        ' A fileName was selected. Add the code to save the file here
vbwProfiler.vbwExecuteLine 298
        sFileName = oCDLG.cFileName(0)

        'save the path
vbwProfiler.vbwExecuteLine 299
        If sType = "Text" Then
vbwProfiler.vbwExecuteLine 300
            Settings.TextExportPath = ExtractPathFromFile(sFileName)
        Else
vbwProfiler.vbwExecuteLine 301 'B
vbwProfiler.vbwExecuteLine 302
            Settings.HTMLExportPath = ExtractPathFromFile(sFileName)
        End If
vbwProfiler.vbwExecuteLine 303 'B
vbwProfiler.vbwExecuteLine 304
        DoEvents
    Else
vbwProfiler.vbwExecuteLine 305 'B
vbwProfiler.vbwProcOut 12
vbwProfiler.vbwExecuteLine 306
        Exit Sub
    End If
vbwProfiler.vbwExecuteLine 307 'B

    ' now generate the actual file
vbwProfiler.vbwExecuteLine 308
    Open sFileName For Output As #2
vbwProfiler.vbwExecuteLine 309
    sTemp = createGURPSText(sType)
vbwProfiler.vbwExecuteLine 310
    Print #2, sTemp
vbwProfiler.vbwExecuteLine 311
    Close #2

    Dim retval As Long
    Dim sProgramPath As String

    'make sure the path is set
vbwProfiler.vbwExecuteLine 312
    If sType = "Text" Or sType = "Text Slim" Then
vbwProfiler.vbwExecuteLine 313
        sProgramPath = Settings.TextViewerPath
    Else
vbwProfiler.vbwExecuteLine 314 'B
vbwProfiler.vbwExecuteLine 315
        sProgramPath = Settings.HTMLBrowserPath
    End If
vbwProfiler.vbwExecuteLine 316 'B

vbwProfiler.vbwExecuteLine 317
    If sProgramPath = "" Then
vbwProfiler.vbwExecuteLine 318
        MsgBox "No viewer specified."
vbwProfiler.vbwExecuteLine 319
        frmConfigure.Show vbModal, frmDesigner
vbwProfiler.vbwExecuteLine 320
        Set frmConfigure = Nothing

    Else 'attempt to launch the program
vbwProfiler.vbwExecuteLine 321 'B
vbwProfiler.vbwExecuteLine 322
        retval = StartDoc(sProgramPath)
vbwProfiler.vbwExecuteLine 323
        If retval <= 32 Then        ' Error
vbwProfiler.vbwExecuteLine 324
            MsgBox "Web Page not Opened", vbExclamation, "URL Failed"
        End If
vbwProfiler.vbwExecuteLine 325 'B
    End If
vbwProfiler.vbwExecuteLine 326 'B
vbwProfiler.vbwProcOut 12
vbwProfiler.vbwExecuteLine 327
    Exit Sub

' check to see if the user Cancels out instead of electing to save a file
errorhandler:
vbwProfiler.vbwExecuteLine 328
    InfoPrint 1, "Error in ExportFile:  " + CStr(err.Number) + " " + err.Description
vbwProfiler.vbwExecuteLine 329
    Resume Next
vbwProfiler.vbwProcOut 12
vbwProfiler.vbwExecuteLine 330
End Sub

Function LoadRecords(ByVal FileName As String) As Boolean
' This function just loads the file data into an array of
' "components".  They are not actually added to the "Vehicle" at this point
' that is done in the RebuildComponentStructure function which is called at the end of this sub
' if the vehicle data was able to be read in properly
vbwProfiler.vbwProcIn 13

Dim sVersInfo() As String
vbwProfiler.vbwExecuteLine 331
Const Major = 0
vbwProfiler.vbwExecuteLine 332
Const Minor = 1
vbwProfiler.vbwExecuteLine 333
Const Revision = 2
vbwProfiler.vbwExecuteLine 334
Const REG_ID = 3

Dim iFreeFile As Long
Dim sLine As String


vbwProfiler.vbwExecuteLine 335
On Error GoTo errorhandler

' Destroy any old vehicle object and create new
vbwProfiler.vbwExecuteLine 336
Set Vehicle = Nothing
vbwProfiler.vbwExecuteLine 337
Set Vehicle = New Vehicles.clsVehicle
'todo: MUST use obtptr() of IComponent interface of Vehicle for Key in tree for Vehicle
' Clear the treeview of any nodes that might already exist
vbwProfiler.vbwExecuteLine 338
frmDesigner.treeVehicle.Nodes.Clear
vbwProfiler.vbwExecuteLine 339
z = m_oCurrentVeh.GetOverDrive

'//show the "loading" in the status bar
vbwProfiler.vbwExecuteLine 340
With frmDesigner
vbwProfiler.vbwExecuteLine 341
    .MousePointer = vbHourglass
vbwProfiler.vbwExecuteLine 342
    .ListView1.MousePointer = vbHourglass
vbwProfiler.vbwExecuteLine 343
End With

' set the status bar panels
vbwProfiler.vbwExecuteLine 344
frmDesigner.StatusBar1.Panels(1).Text = "Reading file data  0%"
vbwProfiler.vbwExecuteLine 345
frmDesigner.StatusBar1.Panels(1).Picture = frmDesigner.ImageList1.ListImages(11).Picture

' determine whether we are dealing with a old file format or new
vbwProfiler.vbwExecuteLine 346
If NewFileFormat(FileName) Then
vbwProfiler.vbwExecuteLine 347
    If LoadComponents_NewFormat(FileName) = False Then
vbwProfiler.vbwExecuteLine 348
         GoTo errorhandler
    End If
vbwProfiler.vbwExecuteLine 349 'B
vbwProfiler.vbwExecuteLine 350
    Debug.Print "LoadRecords: GVD User Registration ID = " & gsRegID  'MPJ 07/04/2000
Else
vbwProfiler.vbwExecuteLine 351 'B
    'open the file
vbwProfiler.vbwExecuteLine 352
    iFreeFile = FreeFile
vbwProfiler.vbwExecuteLine 353
    Open FileName For Input As #iFreeFile
    'get the first line which has the version info
vbwProfiler.vbwExecuteLine 354
    Line Input #iFreeFile, sLine
vbwProfiler.vbwExecuteLine 355
    Close #iFreeFile

vbwProfiler.vbwExecuteLine 356
    sLine = DecryptINI(DecryptINI$(sLine, z), z & Str(5982))  'the first line is double encrypted
vbwProfiler.vbwExecuteLine 357
    sVersInfo = Split(sLine, ",")
vbwProfiler.vbwExecuteLine 358
    gsMajor = sVersInfo(Major)
vbwProfiler.vbwExecuteLine 359
    gsMinor = sVersInfo(Minor)
vbwProfiler.vbwExecuteLine 360
    gsRevision = sVersInfo(Revision)

vbwProfiler.vbwExecuteLine 361
    If LoadComponents_OldFormat(FileName) = False Then
vbwProfiler.vbwExecuteLine 362
         GoTo errorhandler
    End If
vbwProfiler.vbwExecuteLine 363 'B
End If
vbwProfiler.vbwExecuteLine 364 'B

'//reset the status bar panels
vbwProfiler.vbwExecuteLine 365
frmDesigner.StatusBar1.Panels(1).Text = ""
vbwProfiler.vbwExecuteLine 366
frmDesigner.StatusBar1.Panels(1).Picture = LoadPicture()

'//Rebuild Vehicle Structure
vbwProfiler.vbwExecuteLine 367
If RebuildComponentStructure = False Then
vbwProfiler.vbwExecuteLine 368
     GoTo errorhandler
End If
vbwProfiler.vbwExecuteLine 369 'B

vbwProfiler.vbwExecuteLine 370
With frmDesigner
vbwProfiler.vbwExecuteLine 371
    .MousePointer = vbDefault
vbwProfiler.vbwExecuteLine 372
    .ListView1.MousePointer = vbDefault
vbwProfiler.vbwExecuteLine 373
    .treeVehicle.Nodes(BODY_KEY).Expanded = True
vbwProfiler.vbwExecuteLine 374
End With

'//load was successful
vbwProfiler.vbwExecuteLine 375
LoadRecords = True
vbwProfiler.vbwProcOut 13
vbwProfiler.vbwExecuteLine 376
Exit Function

errorhandler:
vbwProfiler.vbwExecuteLine 377
    With frmDesigner
vbwProfiler.vbwExecuteLine 378
        .MousePointer = vbDefault
vbwProfiler.vbwExecuteLine 379
        .ListView1.MousePointer = vbDefault
vbwProfiler.vbwExecuteLine 380
    End With

vbwProfiler.vbwExecuteLine 381
    MsgBox "Unable to load file.  Vehicle file is either invalid or corrupt."
vbwProfiler.vbwExecuteLine 382
    LoadRecords = False
vbwProfiler.vbwExecuteLine 383
    Close #iFreeFile
vbwProfiler.vbwProcOut 13
vbwProfiler.vbwExecuteLine 384
End Function

Function NewFileFormat(sFileName As String) As Boolean
vbwProfiler.vbwProcIn 14

    Dim iFree As Long
    Dim bSig As Byte
    Dim uHeader As Header
    Dim uHeader2 As Header2

vbwProfiler.vbwExecuteLine 385
    On Error GoTo errorhandler

vbwProfiler.vbwExecuteLine 386
    iFree = FreeFile
vbwProfiler.vbwExecuteLine 387
    Open sFileName For Binary As #iFree

vbwProfiler.vbwExecuteLine 388
    Get #iFree, , bSig

    ' returns True if bSig matches our Signature constant
vbwProfiler.vbwExecuteLine 389
    If (bSig <> SIG_128) And (bSig <> SIG_129) Then
vbwProfiler.vbwExecuteLine 390
        NewFileFormat = False
    Else
vbwProfiler.vbwExecuteLine 391 'B
        'get the rest of the header
vbwProfiler.vbwExecuteLine 392
        Get #iFree, 2, uHeader
vbwProfiler.vbwExecuteLine 393
        With uHeader
vbwProfiler.vbwExecuteLine 394
            gsMajor = CStr(.Major)
vbwProfiler.vbwExecuteLine 395
            gsMinor = CStr(.Minor)
vbwProfiler.vbwExecuteLine 396
            gsRevision = CStr(.Revision)
vbwProfiler.vbwExecuteLine 397
        End With

vbwProfiler.vbwExecuteLine 398
        NewFileFormat = True

        ' set the appropriate offset for where our actual Vehicle data starts
vbwProfiler.vbwExecuteLine 399
        If bSig = SIG_129 Then
vbwProfiler.vbwExecuteLine 400
            m_lngOffset = OFFSET2

            ' get the second Header of the new file format so we can
            'obtain the GUID
vbwProfiler.vbwExecuteLine 401
            Get #iFree, OFFSET1, uHeader2
vbwProfiler.vbwExecuteLine 402
            p_sGUID = uHeader2.GUID

'vbwLine 403:        ElseIf bSig = SIG_128 Then
        ElseIf vbwProfiler.vbwExecuteLine(403) Or bSig = SIG_128 Then
vbwProfiler.vbwExecuteLine 404
            m_lngOffset = OFFSET1
        Else
vbwProfiler.vbwExecuteLine 405 'B
vbwProfiler.vbwExecuteLine 406
            NewFileFormat = False
        End If
vbwProfiler.vbwExecuteLine 407 'B
    End If
vbwProfiler.vbwExecuteLine 408 'B

vbwProfiler.vbwExecuteLine 409
    Close #iFree
vbwProfiler.vbwProcOut 14
vbwProfiler.vbwExecuteLine 410
    Exit Function

errorhandler:
vbwProfiler.vbwExecuteLine 411
    NewFileFormat = False
vbwProfiler.vbwExecuteLine 412
    Close #iFree
vbwProfiler.vbwProcOut 14
vbwProfiler.vbwExecuteLine 413
End Function


Function LoadComponents_NewFormat(ByVal sFileName As String) As Boolean
     '//accepts a filename and loads in all the components
vbwProfiler.vbwProcIn 15
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
vbwProfiler.vbwExecuteLine 414
    Set oZip = New cZlib



vbwProfiler.vbwExecuteLine 415
    On Error GoTo errorhandler

vbwProfiler.vbwExecuteLine 416
    ReDim uRet(1)
vbwProfiler.vbwExecuteLine 417
    iFree = FreeFile


vbwProfiler.vbwExecuteLine 418
    Open sFileName For Binary As iFree

    'determine the length of the data
vbwProfiler.vbwExecuteLine 419
    lngDataLen = FileLen(sFileName) - m_lngOffset
vbwProfiler.vbwExecuteLine 420
    ReDim b(0 To lngDataLen - 1)

    'read it all in and decompress it
vbwProfiler.vbwExecuteLine 421
    Get #iFree, m_lngOffset, b

vbwProfiler.vbwExecuteLine 422
    If FLAG_NOZIP <> True Then
vbwProfiler.vbwExecuteLine 423
        oZip.UncompressB b
    End If
vbwProfiler.vbwExecuteLine 424 'B

    ' convert to string and split this up into our seperate component lines
vbwProfiler.vbwExecuteLine 425
    sTemp = StrConv(b, vbUnicode)

vbwProfiler.vbwExecuteLine 426
    sArray = Split(sTemp, Chr(254))


vbwProfiler.vbwExecuteLine 427
   k = 0
'vbwLine 428:    Do While k <= UBound(sArray)
    Do While vbwProfiler.vbwExecuteLine(428) Or k <= UBound(sArray)

        'fill the variant array that we will be passing into the FileLoader class
vbwProfiler.vbwExecuteLine 429
        If Left(sArray(k), 1) = "[" Then
            'now remove the leading and trailing characters
vbwProfiler.vbwExecuteLine 430
            i = i + 1
vbwProfiler.vbwExecuteLine 431
            j = 0
vbwProfiler.vbwExecuteLine 432
            ReDim Preserve Components(i)
vbwProfiler.vbwExecuteLine 433
            sArray(k) = Mid(sArray(k), 2, Len(sArray(k)) - 2)
vbwProfiler.vbwExecuteLine 434
            Components(i).TreeInfo = sArray(k)

        Else
vbwProfiler.vbwExecuteLine 435 'B
vbwProfiler.vbwExecuteLine 436
            j = j + 1
vbwProfiler.vbwExecuteLine 437
            ReDim Preserve Components(i).Properties(j)
vbwProfiler.vbwExecuteLine 438
            Components(i).Properties(j) = sArray(k)
        End If
vbwProfiler.vbwExecuteLine 439 'B
vbwProfiler.vbwExecuteLine 440
        k = k + 1
vbwProfiler.vbwExecuteLine 441
        frmDesigner.StatusBar1.Panels(1).Text = "Reading file data  " & CLng(k / UBound(sArray) * 100) & "%"
vbwProfiler.vbwExecuteLine 442
    Loop


vbwProfiler.vbwExecuteLine 443
    Close #iFree
vbwProfiler.vbwExecuteLine 444
    Set oZip = Nothing
vbwProfiler.vbwExecuteLine 445
    LoadComponents_NewFormat = True
vbwProfiler.vbwExecuteLine 446
    Debug.Print "LoadComponents_NewFormat: " & sTemp
vbwProfiler.vbwProcOut 15
vbwProfiler.vbwExecuteLine 447
    Exit Function

errorhandler:
vbwProfiler.vbwExecuteLine 448
    Debug.Print "LoadComponents_NewFormat: " & err.Description
vbwProfiler.vbwExecuteLine 449
    Set oZip = Nothing
vbwProfiler.vbwExecuteLine 450
    LoadComponents_NewFormat = False
vbwProfiler.vbwProcOut 15
vbwProfiler.vbwExecuteLine 451
End Function

Function LoadComponents_OldFormat(ByVal sFileName As String) As Boolean
    '//accepts a filename and loads in all the components
vbwProfiler.vbwProcIn 16
    Dim iFree As Long
    Dim sTemp As String
    Dim i As Long
    Dim j As Long
    Dim lngFileLen As Long
    Dim lngBytesRead As Long

vbwProfiler.vbwExecuteLine 452
    On Error GoTo errorhandler

vbwProfiler.vbwExecuteLine 453
    ReDim uRet(1)
vbwProfiler.vbwExecuteLine 454
    iFree = FreeFile

vbwProfiler.vbwExecuteLine 455
    lngFileLen = FileLen(sFileName)
vbwProfiler.vbwExecuteLine 456
    frmDesigner.StatusBar1.Panels(1).Text = "Reading file data  0%"
vbwProfiler.vbwExecuteLine 457
    frmDesigner.StatusBar1.Panels(1).Picture = frmDesigner.ImageList1.ListImages(11).Picture

vbwProfiler.vbwExecuteLine 458
    Open sFileName For Input As iFree

    'load the first line and skip it
vbwProfiler.vbwExecuteLine 459
    Line Input #iFree, sTemp


'vbwLine 460:    Do While Not EOF(iFree)
    Do While vbwProfiler.vbwExecuteLine(460) Or Not EOF(iFree)
vbwProfiler.vbwExecuteLine 461
        Line Input #iFree, sTemp
vbwProfiler.vbwExecuteLine 462
        lngBytesRead = lngBytesRead + Len(sTemp)
vbwProfiler.vbwExecuteLine 463
        sTemp = DecryptINI$(sTemp, z)
        'fill the variant array that we will be passing into the FileLoader class
vbwProfiler.vbwExecuteLine 464
        If Left(sTemp, 1) = "[" Then
            'now remove the leading and trailing characters
vbwProfiler.vbwExecuteLine 465
            i = i + 1
vbwProfiler.vbwExecuteLine 466
            j = 0
vbwProfiler.vbwExecuteLine 467
            ReDim Preserve Components(i)
vbwProfiler.vbwExecuteLine 468
            sTemp = Mid(sTemp, 2, Len(sTemp) - 2)
vbwProfiler.vbwExecuteLine 469
            Components(i).TreeInfo = sTemp

        Else
vbwProfiler.vbwExecuteLine 470 'B
vbwProfiler.vbwExecuteLine 471
            j = j + 1
vbwProfiler.vbwExecuteLine 472
            ReDim Preserve Components(i).Properties(j)
vbwProfiler.vbwExecuteLine 473
            Components(i).Properties(j) = sTemp
        End If
vbwProfiler.vbwExecuteLine 474 'B
vbwProfiler.vbwExecuteLine 475
        frmDesigner.StatusBar1.Panels(1).Text = "Reading file data  " & CLng(lngBytesRead / lngFileLen * 100) & "%"
vbwProfiler.vbwExecuteLine 476
    Loop


vbwProfiler.vbwExecuteLine 477
    Close #iFree

vbwProfiler.vbwExecuteLine 478
    LoadComponents_OldFormat = True
vbwProfiler.vbwProcOut 16
vbwProfiler.vbwExecuteLine 479
    Exit Function

errorhandler:

vbwProfiler.vbwExecuteLine 480
    LoadComponents_OldFormat = False

vbwProfiler.vbwProcOut 16
vbwProfiler.vbwExecuteLine 481
End Function

Function RebuildComponentStructure() As Boolean
    ' This is where the read in saved vehicle data is turned into the actual Vehicle heirarchy.
    ' It makes calls to Vehicle.AddObject for creating the correct objects based on the parsed datatypes
vbwProfiler.vbwProcIn 17
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

vbwProfiler.vbwExecuteLine 482
    On Error GoTo errorhandler
    '/show our progress meter
vbwProfiler.vbwExecuteLine 483
    lngUpper = UBound(Components)
vbwProfiler.vbwExecuteLine 484
    frmDesigner.StatusBar1.Panels(1).Text = "Building tree  0%"

    ' create an instance of the FileLoader class
vbwProfiler.vbwExecuteLine 485
    Set tobj = New clsFileLoader

    '//load up our tree and Component object
vbwProfiler.vbwExecuteLine 486
    For iCount = 1 To lngUpper
vbwProfiler.vbwExecuteLine 487
        vc = Split(Components(iCount).TreeInfo, "|")
vbwProfiler.vbwExecuteLine 488
        sDescription = vc(0)
vbwProfiler.vbwExecuteLine 489
        sKey = vc(1)
vbwProfiler.vbwExecuteLine 490
        sParent = vc(2)
vbwProfiler.vbwExecuteLine 491
        dType = vc(3)
vbwProfiler.vbwExecuteLine 492
        icon1 = vc(4)

         'create the vehicle component object
vbwProfiler.vbwExecuteLine 493
         If m_oCurrentVeh.addObject(dType, sKey, sParent, icon1, sDescription, True) Then
             'create the tree node (unless its like a weapon link, performance, etc)
vbwProfiler.vbwExecuteLine 494
             With frmDesigner.treeVehicle
vbwProfiler.vbwExecuteLine 495
                If sKey = BODY_KEY Then 'if its the body, then its the root node and doesnt have a parent
vbwProfiler.vbwExecuteLine 496
                    .Nodes.Add , , sKey, sDescription, icon1
'vbwLine 497:                ElseIf (dType = PERFORMANCEPROFILE) Or (dType = WeaponLink) Then
                ElseIf vbwProfiler.vbwExecuteLine(497) Or (dType = PERFORMANCEPROFILE) Or (dType = WeaponLink) Then
                    'performance profiles and weapon links do NOT get added to the tree
                    'todo: They will have to now!
                Else
vbwProfiler.vbwExecuteLine 498 'B
vbwProfiler.vbwExecuteLine 499
                    .Nodes.Add sParent, tvwChild, sKey, sDescription, icon1
                End If
vbwProfiler.vbwExecuteLine 500 'B
vbwProfiler.vbwExecuteLine 501
            End With
        Else
vbwProfiler.vbwExecuteLine 502 'B
vbwProfiler.vbwExecuteLine 503
            GoTo errorhandler
        End If
vbwProfiler.vbwExecuteLine 504 'B

vbwProfiler.vbwExecuteLine 505
        For iPropCount = 1 To UBound(Components(iCount).Properties)
vbwProfiler.vbwExecuteLine 506
            vc = Split(Components(iCount).Properties(iPropCount), "|")
            'check for keychain.
vbwProfiler.vbwExecuteLine 507
            If vc(1) = "[" Then

vbwProfiler.vbwExecuteLine 508
                j = 1
vbwProfiler.vbwExecuteLine 509
                ReDim arrkey(UBound(vc) - 1)
vbwProfiler.vbwExecuteLine 510
                For i = 2 To UBound(vc)
vbwProfiler.vbwExecuteLine 511
                    arrkey(j) = vc(i)
vbwProfiler.vbwExecuteLine 512
                    j = j + 1
vbwProfiler.vbwExecuteLine 513
                Next
vbwProfiler.vbwExecuteLine 514
                propvalue = arrkey
vbwProfiler.vbwExecuteLine 515
                memberID = vc(0)
vbwProfiler.vbwExecuteLine 516
                tobj.LetProperties m_oCurrentVeh.Components(sKey), memberID, propvalue
            Else
vbwProfiler.vbwExecuteLine 517 'B
                'fill the properties for this object
vbwProfiler.vbwExecuteLine 518
                Debug.Assert vc(0) <> "CombinedComponentVolume"
vbwProfiler.vbwExecuteLine 519
                memberID = vc(0)
vbwProfiler.vbwExecuteLine 520
                propvalue = vc(1)
vbwProfiler.vbwExecuteLine 521
                tobj.LetProperties m_oCurrentVeh.Components(sKey), memberID, propvalue
            End If
vbwProfiler.vbwExecuteLine 522 'B
vbwProfiler.vbwExecuteLine 523
        Next

vbwProfiler.vbwExecuteLine 524
        frmDesigner.StatusBar1.Panels(1).Text = "Building tree  " & CLng(iCount / lngUpper * 100) & "%"
vbwProfiler.vbwExecuteLine 525
    Next

vbwProfiler.vbwExecuteLine 526
    RebuildComponentStructure = True
    'destroy the instance of the fileloader class
vbwProfiler.vbwExecuteLine 527
    Set tobj = Nothing
vbwProfiler.vbwExecuteLine 528
    frmDesigner.StatusBar1.Panels(1).Picture = LoadPicture()
vbwProfiler.vbwExecuteLine 529
    frmDesigner.StatusBar1.Panels(1).Text = ""
vbwProfiler.vbwProcOut 17
vbwProfiler.vbwExecuteLine 530
    Exit Function

errorhandler:
vbwProfiler.vbwExecuteLine 531
    Debug.Print "RebuildComponentStructure: " & err.Description
vbwProfiler.vbwExecuteLine 532
    frmDesigner.StatusBar1.Panels(1).Text = ""
vbwProfiler.vbwExecuteLine 533
    frmDesigner.StatusBar1.Panels(1).Picture = LoadPicture()
vbwProfiler.vbwExecuteLine 534
    RebuildComponentStructure = False

vbwProfiler.vbwProcOut 17
vbwProfiler.vbwExecuteLine 535
End Function
'//////////////////////////////////////////////////////
Function CreatePrintString(ByVal vc As Variant) As String()
vbwProfiler.vbwProcIn 18

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
vbwProfiler.vbwExecuteLine 536
For i = 1 To UBound(vc)
vbwProfiler.vbwExecuteLine 537
    If (vc(i, 0) = "Datatype") Or (vc(i, 0) = "datatype") Then
vbwProfiler.vbwExecuteLine 538
        dType = vc(0, i)
'vbwLine 539:    ElseIf (vc(i, 0) = "Key") Or (vc(i, 0) = "key") Then
    ElseIf vbwProfiler.vbwExecuteLine(539) Or (vc(i, 0) = "Key") Or (vc(i, 0) = "key") Then
vbwProfiler.vbwExecuteLine 540
        sKey = vc(0, i)
vbwProfiler.vbwExecuteLine 541
        If sKey = BODY_KEY Then
vbwProfiler.vbwExecuteLine 542
             sParent = "0_"
        End If
vbwProfiler.vbwExecuteLine 543 'B
'vbwLine 544:    ElseIf (vc(i, 0) = "Parent") Or (vc(i, 0) = "parent") Then
    ElseIf vbwProfiler.vbwExecuteLine(544) Or (vc(i, 0) = "Parent") Or (vc(i, 0) = "parent") Then
vbwProfiler.vbwExecuteLine 545
        sParent = vc(0, i)
'vbwLine 546:    ElseIf (vc(i, 0) = "customdescription") Or (vc(i, 0) = "Customdescription") Or (vc(i, 0) = "CustomDescription") Then
    ElseIf vbwProfiler.vbwExecuteLine(546) Or (vc(i, 0) = "customdescription") Or (vc(i, 0) = "Customdescription") Or (vc(i, 0) = "CustomDescription") Then
vbwProfiler.vbwExecuteLine 547
        sDescription = vc(0, i)
'vbwLine 548:    ElseIf (vc(i, 0) = "Image") Or (vc(i, 0) = "image") Then
    ElseIf vbwProfiler.vbwExecuteLine(548) Or (vc(i, 0) = "Image") Or (vc(i, 0) = "image") Then
vbwProfiler.vbwExecuteLine 549
        icon1 = vc(0, i)
'vbwLine 550:    ElseIf (vc(i, 0) = "SelectedImage") Or (vc(i, 0) = "selectedimage") Or (vc(i, 0) = "Selectedimage") Then
    ElseIf vbwProfiler.vbwExecuteLine(550) Or (vc(i, 0) = "SelectedImage") Or (vc(i, 0) = "selectedimage") Or (vc(i, 0) = "Selectedimage") Then
vbwProfiler.vbwExecuteLine 551
        icon2 = vc(0, i)
    End If
vbwProfiler.vbwExecuteLine 552 'B
vbwProfiler.vbwExecuteLine 553
    If (dType <> vbEmpty) And (sKey <> "") And (sDescription <> "") And (icon1 <> vbEmpty) And (icon2 <> vbEmpty) And (sParent <> "") Then
vbwProfiler.vbwExecuteLine 554
         Exit For
    End If
vbwProfiler.vbwExecuteLine 555 'B
vbwProfiler.vbwExecuteLine 556
Next

'store this line in the array
vbwProfiler.vbwExecuteLine 557
ReDim retval(1)
vbwProfiler.vbwExecuteLine 558
retval(1) = "[" + sDescription + "|" + sKey + "|" + sParent + "|" + Str(dType) + "|" + Str(icon1) + "|" + Str(icon2) + "]"

vbwProfiler.vbwExecuteLine 559
ReDim Preserve retval(UBound(vc) + 1)
vbwProfiler.vbwExecuteLine 560
SIZE = 1

vbwProfiler.vbwExecuteLine 561
For i = 1 To UBound(vc)
vbwProfiler.vbwExecuteLine 562
    SIZE = SIZE + 1
vbwProfiler.vbwExecuteLine 563
    vType = VarType(vc(0, i))
vbwProfiler.vbwExecuteLine 564
    svType = Str(vType)
    'check for keychains (vbarray + vbvariant)
    'reset the skeychain string
vbwProfiler.vbwExecuteLine 565
    skeychain = ""
vbwProfiler.vbwExecuteLine 566
    If vType = vbArray + vbVariant Then
        'found a keychain.  Place the "[" which inidcates a keychain
vbwProfiler.vbwExecuteLine 567
        skeychain = skeychain + "|" + "["
        'Seperate the individual keys into a string
vbwProfiler.vbwExecuteLine 568
        For j = 1 To UBound(vc(0, i))

vbwProfiler.vbwExecuteLine 569
            skeychain = skeychain + "|" + vc(0, i)(j)
vbwProfiler.vbwExecuteLine 570
        Next
vbwProfiler.vbwExecuteLine 571
        retval(SIZE) = vc(i, 0) + skeychain
'vbwLine 572:    ElseIf vType = vbArray + vbString Then
    ElseIf vbwProfiler.vbwExecuteLine(572) Or vType = vbArray + vbString Then
        'found a keychain.  Place the "[" which inidcates a keychain
vbwProfiler.vbwExecuteLine 573
        skeychain = skeychain + "|" + "["
        'Seperate the individual keys into a string
vbwProfiler.vbwExecuteLine 574
        For j = 1 To UBound(vc(0, i))

vbwProfiler.vbwExecuteLine 575
            skeychain = skeychain + "|" + vc(0, i)(j)
vbwProfiler.vbwExecuteLine 576
        Next
vbwProfiler.vbwExecuteLine 577
        retval(SIZE) = vc(i, 0) + skeychain
'vbwLine 578:    ElseIf vType <> vbString Then
    ElseIf vbwProfiler.vbwExecuteLine(578) Or vType <> vbString Then
vbwProfiler.vbwExecuteLine 579
        retval(SIZE) = vc(i, 0) + "|" + Str(vc(0, i))
    Else
vbwProfiler.vbwExecuteLine 580 'B
vbwProfiler.vbwExecuteLine 581
        retval(SIZE) = vc(i, 0) + "|" + vc(0, i)
    End If
vbwProfiler.vbwExecuteLine 582 'B
vbwProfiler.vbwExecuteLine 583
Next


vbwProfiler.vbwExecuteLine 584
CreatePrintString = retval
vbwProfiler.vbwProcOut 18
vbwProfiler.vbwExecuteLine 585
End Function

Sub CreateRecords()
vbwProfiler.vbwProcIn 19

    Dim tobj As clsFileLoader
    Dim vc As Variant
    Dim iIndex As Integer
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim keychainkeys As Variant
    Dim sKey As String

vbwProfiler.vbwExecuteLine 586
    On Error GoTo errorhandler

    ' create an instance of the FileLoader class
vbwProfiler.vbwExecuteLine 587
    Set tobj = New clsFileLoader
vbwProfiler.vbwExecuteLine 588
    z = m_oCurrentVeh.GetOverDrive

    'todo: Here instead of using GetFirstParent, we will start with the Body node and only
    ' itterate thru all children/sub children from this node.
    'NOTE: We do have to itterate thru the tree and not simply itterate thru the
    ' m_oCurrentVeh.components collection.  Itterating thru the components collection would result in
    'components being read and potentially the parent to which they must be added not being in the tree yet.
    'Proof: Lets say you add a Weapon to the Body and then a turret to the body.  Now lets say you move the weapon
    ' to the turret.  That puts the weapon ahead of the Turret in the collection so when reading in the weapon
    ' it would fail when trying to addobject to the parent turret which hasnt been installed yet.
vbwProfiler.vbwExecuteLine 589
    GetFirstParent 'Find a root node in the treeview
    'get the index of the root node that is at the top of the treeview
vbwProfiler.vbwExecuteLine 590
    iIndex = frmDesigner.treeVehicle.Nodes(p_nIndex).FirstSibling.index
vbwProfiler.vbwExecuteLine 591
    sKey = frmDesigner.treeVehicle.Nodes(iIndex).Key
    ' debug I really dont need a select case here since the first node is always the Body
    'sName = TypeName(m_oCurrentVeh.Components.item(sKey))
vbwProfiler.vbwExecuteLine 592
    vc = tobj.GetProperties(Vehicle.Components(sKey))
vbwProfiler.vbwExecuteLine 593
    i = 1
vbwProfiler.vbwExecuteLine 594
    ReDim Components(i)
vbwProfiler.vbwExecuteLine 595
    Components(1).Properties = CreatePrintString(vc)

    'If the Node has Children call the sub that writes the children
vbwProfiler.vbwExecuteLine 596
    If frmDesigner.treeVehicle.Nodes(iIndex).children > 0 Then
vbwProfiler.vbwExecuteLine 597
        WriteChild iIndex
    End If
vbwProfiler.vbwExecuteLine 598 'B

    'Now save the Performance Profiles which are not
    'visually represented by the Tree
vbwProfiler.vbwExecuteLine 599
    keychainkeys = m_oCurrentVeh.Components(BODY_KEY).PerformanceProfileKeychain
vbwProfiler.vbwExecuteLine 600
    If (UBound(keychainkeys) >= 1) And keychainkeys(1) <> "" Then
vbwProfiler.vbwExecuteLine 601
        For k = 1 To UBound(keychainkeys)
vbwProfiler.vbwExecuteLine 602
            vc = tobj.GetProperties(m_oCurrentVeh.Components(keychainkeys(k)))
vbwProfiler.vbwExecuteLine 603
            i = UBound(Components) + 1
vbwProfiler.vbwExecuteLine 604
            ReDim Preserve Components(i)
vbwProfiler.vbwExecuteLine 605
            Components(i).Properties = CreatePrintString(vc)
vbwProfiler.vbwExecuteLine 606
        Next
    End If
vbwProfiler.vbwExecuteLine 607 'B
    'Now save the Weapon Links which are not
    'visually represented by the Tree
vbwProfiler.vbwExecuteLine 608
    keychainkeys = m_oCurrentVeh.Components(BODY_KEY).WeaponLinkKeychain
vbwProfiler.vbwExecuteLine 609
    If (UBound(keychainkeys) >= 1) And keychainkeys(1) <> "" Then
vbwProfiler.vbwExecuteLine 610
        For k = 1 To UBound(keychainkeys)
vbwProfiler.vbwExecuteLine 611
            vc = tobj.GetProperties(m_oCurrentVeh.Components(keychainkeys(k)))
vbwProfiler.vbwExecuteLine 612
            i = UBound(Components) + 1
vbwProfiler.vbwExecuteLine 613
            ReDim Preserve Components(i)
vbwProfiler.vbwExecuteLine 614
            Components(i).Properties = CreatePrintString(vc)
vbwProfiler.vbwExecuteLine 615
        Next
    End If
vbwProfiler.vbwExecuteLine 616 'B

    'destroy the instance of the fileloader class
vbwProfiler.vbwExecuteLine 617
    Set tobj = Nothing
vbwProfiler.vbwProcOut 19
vbwProfiler.vbwExecuteLine 618
    Exit Sub

errorhandler:


vbwProfiler.vbwProcOut 19
vbwProfiler.vbwExecuteLine 619
End Sub
Sub WriteRecord(ByVal FileName As String)
    'now get ready to print the properties for this object
'first delete the contents  of the file
vbwProfiler.vbwProcIn 20
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

vbwProfiler.vbwExecuteLine 620
iFree = FreeFile

'//reg check
vbwProfiler.vbwExecuteLine 621
tempbyte = ChopCheck
vbwProfiler.vbwExecuteLine 622
If (IsEmpty(tempbyte) = False) And (UBound(tempbyte) - LBound(tempbyte) = UBound(gsRegNum) - LBound(gsRegNum)) Then
vbwProfiler.vbwExecuteLine 623
    For i = 1 To UBound(gsRegNum)
vbwProfiler.vbwExecuteLine 624
        If tempbyte(i) = gsRegNum(i) Then
vbwProfiler.vbwExecuteLine 625
            bFlag = True
        Else
vbwProfiler.vbwExecuteLine 626 'B
vbwProfiler.vbwExecuteLine 627
            bFlag = False
vbwProfiler.vbwExecuteLine 628
            GoTo reghandler
        End If
vbwProfiler.vbwExecuteLine 629 'B
vbwProfiler.vbwExecuteLine 630
    Next
Else
vbwProfiler.vbwExecuteLine 631 'B
vbwProfiler.vbwExecuteLine 632
    bFlag = False
vbwProfiler.vbwExecuteLine 633
    GoTo reghandler
End If
vbwProfiler.vbwExecuteLine 634 'B

vbwProfiler.vbwExecuteLine 635
frmDesigner.StatusBar1.Panels(1).Text = "Reading vehicle data..."
vbwProfiler.vbwExecuteLine 636
CreateRecords

'delete the existing file
vbwProfiler.vbwExecuteLine 637
Open FileName For Random As #iFree
vbwProfiler.vbwExecuteLine 638
Close #iFree

vbwProfiler.vbwExecuteLine 639
frmDesigner.StatusBar1.Panels(1).Text = "Writing file  0%"
'open the file for writing
vbwProfiler.vbwExecuteLine 640
Open FileName For Binary As #iFree

vbwProfiler.vbwExecuteLine 641
Set oZip = New cZlib

vbwProfiler.vbwExecuteLine 642
lngUpper = UBound(Components)

vbwProfiler.vbwExecuteLine 643
ReDim sJoin(1)

vbwProfiler.vbwExecuteLine 644
For i = 1 To lngUpper
    'Print #iFree, EncryptINI$(Components(i).TreeInfo, z)
    's = s & Components(i).TreeInfo & Chr(254)
vbwProfiler.vbwExecuteLine 645
    If k > 0 Then
vbwProfiler.vbwExecuteLine 646
        ReDim Preserve sJoin(UBound(sJoin) + UBound(Components(i).Properties))
    Else
vbwProfiler.vbwExecuteLine 647 'B
vbwProfiler.vbwExecuteLine 648
        ReDim sJoin(UBound(Components(i).Properties))
    End If
vbwProfiler.vbwExecuteLine 649 'B
vbwProfiler.vbwExecuteLine 650
    For j = 1 To UBound(Components(i).Properties)
vbwProfiler.vbwExecuteLine 651
        k = k + 1
vbwProfiler.vbwExecuteLine 652
        sJoin(k) = Components(i).Properties(j)
        'Print #iFree, EncryptINI$(Components(i).Properties(j), z)

vbwProfiler.vbwExecuteLine 653
    Next

vbwProfiler.vbwExecuteLine 654
    frmDesigner.StatusBar1.Panels(1).Text = "Writing file  " & CLng(i / lngUpper * 100) & "%"
vbwProfiler.vbwExecuteLine 655
Next


'ReDim Preserve B(Len(B) - 1)

vbwProfiler.vbwExecuteLine 656
s = Join(sJoin, Chr(254))
'remove the last chr(254) from the end
's = Mid(s, 1, Len(s) - 1)
vbwProfiler.vbwExecuteLine 657
b = StrConv(s, vbFromUnicode)


vbwProfiler.vbwExecuteLine 658
If FLAG_NOZIP <> True Then
vbwProfiler.vbwExecuteLine 659
    oZip.CompressB b
End If
vbwProfiler.vbwExecuteLine 660 'B

'create our file header
'first print the header version info
vbwProfiler.vbwExecuteLine 661
With uHeader
vbwProfiler.vbwExecuteLine 662
    .CRC32 = 100 'todo: need to calc the crc first
vbwProfiler.vbwExecuteLine 663
    .Major = App.Major
vbwProfiler.vbwExecuteLine 664
    .Minor = App.Minor
vbwProfiler.vbwExecuteLine 665
    .Revision = App.Revision
vbwProfiler.vbwExecuteLine 666
    .RegID = gsRegID
vbwProfiler.vbwExecuteLine 667
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
vbwProfiler.vbwExecuteLine 668
Put #iFree, , SIG_129  'file version signature to help identify between all the legacy .veh formats
vbwProfiler.vbwExecuteLine 669
Put #iFree, 2, uHeader
vbwProfiler.vbwExecuteLine 670
Put #iFree, OFFSET1, uHeader2

' now put our actual data
vbwProfiler.vbwExecuteLine 671
Put #iFree, OFFSET2, b
vbwProfiler.vbwExecuteLine 672
Set oZip = Nothing

vbwProfiler.vbwExecuteLine 673
Close #iFree


' Writes the vehicle collection to a file specified by the user
reghandler:
vbwProfiler.vbwExecuteLine 674
    Close #iFree
vbwProfiler.vbwProcOut 20
vbwProfiler.vbwExecuteLine 675
    Exit Sub
vbwProfiler.vbwProcOut 20
vbwProfiler.vbwExecuteLine 676
End Sub

Private Sub WriteChild(ByVal iNodeIndex As Integer)
' Write the child nodes to the table. This sub uses recursion
' to loop through the child nodes. It receives the Index of
' the node that has the children
vbwProfiler.vbwProcIn 21
Dim tobj As clsFileLoader
Dim i As Long
Dim iTempIndex As Integer
Dim Temp() As String
Dim vc As Variant
Dim k As Long

' create an instance of the FileLoader class
vbwProfiler.vbwExecuteLine 677
Set tobj = New clsFileLoader

vbwProfiler.vbwExecuteLine 678
iTempIndex = frmDesigner.treeVehicle.Nodes(iNodeIndex).Child.FirstSibling.index
'Loop through all a Parents Child Nodes
vbwProfiler.vbwExecuteLine 679
i = frmDesigner.treeVehicle.Nodes(iNodeIndex).children
vbwProfiler.vbwExecuteLine 680
k = UBound(Components)

vbwProfiler.vbwExecuteLine 681
ReDim Preserve Components(k + i)

vbwProfiler.vbwExecuteLine 682
For i = 1 To frmDesigner.treeVehicle.Nodes(iNodeIndex).children
vbwProfiler.vbwExecuteLine 683
    vc = tobj.GetProperties(m_oCurrentVeh.Components(frmDesigner.treeVehicle.Nodes(iTempIndex).Key))
vbwProfiler.vbwExecuteLine 684
    k = k + 1
vbwProfiler.vbwExecuteLine 685
    Components(k).Properties = CreatePrintString(vc)

    'If the Node we are on has a child call the Sub again
vbwProfiler.vbwExecuteLine 686
    If frmDesigner.treeVehicle.Nodes(iTempIndex).children > 0 Then
vbwProfiler.vbwExecuteLine 687
       WriteChild (iTempIndex)
    End If
vbwProfiler.vbwExecuteLine 688 'B

    'If we are not on the last child move to the next child Node
vbwProfiler.vbwExecuteLine 689
    If i <> frmDesigner.treeVehicle.Nodes(iNodeIndex).children Then
vbwProfiler.vbwExecuteLine 690
       iTempIndex = frmDesigner.treeVehicle.Nodes(iTempIndex).Next.index
    End If
vbwProfiler.vbwExecuteLine 691 'B
vbwProfiler.vbwExecuteLine 692
Next

'destroy the instance of the fileloader class
vbwProfiler.vbwExecuteLine 693
Set tobj = Nothing

vbwProfiler.vbwProcOut 21
vbwProfiler.vbwExecuteLine 694
End Sub

Public Function PasteNode(sSourceKey() As String, sDestinationKey() As String, sKey() As String)
    '//to copy a node, you simply create a new object of the same type as the Source.
vbwProfiler.vbwProcIn 22
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

vbwProfiler.vbwExecuteLine 695
    On Error GoTo errorhandler

vbwProfiler.vbwExecuteLine 696
    For iCount = 1 To UBound(sSourceKey)

vbwProfiler.vbwExecuteLine 697
        With m_oCurrentVeh.Components(sSourceKey(iCount))
vbwProfiler.vbwExecuteLine 698
            dType = .Datatype
vbwProfiler.vbwExecuteLine 699
            sText = .CustomDescription
vbwProfiler.vbwExecuteLine 700
            iImage1 = .Image


vbwProfiler.vbwExecuteLine 701
        End With

        '//first add it to the tree
vbwProfiler.vbwExecuteLine 702
        frmDesigner.treeVehicle.Nodes.Add sDestinationKey(iCount), tvwChild, sKey(iCount), sText, iImage1
        '//now attempt to create the copied node and add it to the vehicle
vbwProfiler.vbwExecuteLine 703
        If m_oCurrentVeh.addObject(dType, sKey(iCount), sDestinationKey(iCount), iImage1, sText, True) = False Then
vbwProfiler.vbwExecuteLine 704
            frmDesigner.treeVehicle.Nodes.Remove (sKey(iCount))
vbwProfiler.vbwExecuteLine 705
            MsgBox "Error copying node. Copy/Paste operation aborted."
vbwProfiler.vbwProcOut 22
vbwProfiler.vbwExecuteLine 706
            Exit Function
        End If
vbwProfiler.vbwExecuteLine 707 'B

        '//restore the key and parent key since it doesnt get properly applied
        '//in a paste operation
vbwProfiler.vbwExecuteLine 708
        m_oCurrentVeh.Components(sKey(iCount)).Parent = sDestinationKey(iCount)
vbwProfiler.vbwExecuteLine 709
        m_oCurrentVeh.Components(sKey(iCount)).Key = sKey(iCount)
vbwProfiler.vbwExecuteLine 710
        m_oCurrentVeh.Components(sKey(iCount)).Datatype = dType

        '//we need to run the location check ourselves since
        '//we are by passing it in the AddObject using the LoadedFlag = TRUE
vbwProfiler.vbwExecuteLine 711
        If m_oCurrentVeh.Components(sKey(iCount)).LocationCheck = False Then
            ' remove the object from the tree since the AddObject failed
vbwProfiler.vbwExecuteLine 712
            frmDesigner.treeVehicle.Nodes.Remove (sKey(iCount))
vbwProfiler.vbwExecuteLine 713
            m_oCurrentVeh.Components.Remove sKey(iCount)
vbwProfiler.vbwExecuteLine 714
            MsgBox "Invalid paste location for this component. Paste aborted."
vbwProfiler.vbwProcOut 22
vbwProfiler.vbwExecuteLine 715
            Exit Function
        End If
vbwProfiler.vbwExecuteLine 716 'B

        Dim sLoc As String
        'get the location so we can restore it
vbwProfiler.vbwExecuteLine 717
        sLoc = m_oCurrentVeh.Components(sKey(iCount)).location

        '//now we need to restore all the values EXCEPT for the Parent and Key values
        '//crap and also power system values
vbwProfiler.vbwExecuteLine 718
         Set tobj = New clsFileLoader
vbwProfiler.vbwExecuteLine 719
         vc = tobj.GetProperties(m_oCurrentVeh.Components(sSourceKey(iCount)))

vbwProfiler.vbwExecuteLine 720
        For i = 1 To UBound(vc)
vbwProfiler.vbwExecuteLine 721
            memberID = vc(i, 0)
vbwProfiler.vbwExecuteLine 722
            propvalue = vc(0, i)
vbwProfiler.vbwExecuteLine 723
            If (memberID = "Location") Or (memberID = "Key") Or (memberID = "Parent") Then
'vbwLine 724:            ElseIf (VarType(propvalue) = vbArray + vbVariant) Or (VarType(propvalue) = vbArray + vbString) Then
            ElseIf vbwProfiler.vbwExecuteLine(724) Or (VarType(propvalue) = vbArray + vbVariant) Or (VarType(propvalue) = vbArray + vbString) Then
            Else
vbwProfiler.vbwExecuteLine 725 'B
vbwProfiler.vbwExecuteLine 726
                tobj.LetProperties m_oCurrentVeh.Components(sKey(iCount)), memberID, propvalue
            End If
vbwProfiler.vbwExecuteLine 727 'B
vbwProfiler.vbwExecuteLine 728
        Next

        'finally, we can add our keychain keys
vbwProfiler.vbwExecuteLine 729
         m_oCurrentVeh.keymanager.AddKeyChainKeys sKey(iCount)
vbwProfiler.vbwExecuteLine 730
    Next

vbwProfiler.vbwProcOut 22
vbwProfiler.vbwExecuteLine 731
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 732
    If err.Number = 35602 Then
        '//we need to change all the keys
vbwProfiler.vbwExecuteLine 733
        sNewKey = GetNextKey
vbwProfiler.vbwExecuteLine 734
        sOldKey = sKey(iCount)
vbwProfiler.vbwExecuteLine 735
        For m = iCount To UBound(sKey)
vbwProfiler.vbwExecuteLine 736
            If sKey(m) = sOldKey Then
vbwProfiler.vbwExecuteLine 737
                sKey(m) = sNewKey
            End If
vbwProfiler.vbwExecuteLine 738 'B
vbwProfiler.vbwExecuteLine 739
        Next
vbwProfiler.vbwExecuteLine 740
        Resume
    End If
vbwProfiler.vbwExecuteLine 741 'B

vbwProfiler.vbwProcOut 22
vbwProfiler.vbwExecuteLine 742
End Function


Function EncryptINI$(ByVal Strg$, ByVal Password$)
vbwProfiler.vbwProcIn 23
   Dim b$, s$, i As Integer, j As Integer
   Dim A1 As Integer, A2 As Integer, A3 As Integer, P$
vbwProfiler.vbwExecuteLine 743
   j = 1
vbwProfiler.vbwExecuteLine 744
   For i = 1 To Len(Password$)
vbwProfiler.vbwExecuteLine 745
     P$ = P$ & Asc(Mid$(Password$, i, 1))
vbwProfiler.vbwExecuteLine 746
   Next

vbwProfiler.vbwExecuteLine 747
   For i = 1 To Len(Strg$)
vbwProfiler.vbwExecuteLine 748
     A1 = Asc(Mid$(P$, j, 1))
vbwProfiler.vbwExecuteLine 749
     j = j + 1
vbwProfiler.vbwExecuteLine 750
     If j > Len(P$) Then
vbwProfiler.vbwExecuteLine 751
          j = 1
     End If
vbwProfiler.vbwExecuteLine 752 'B
vbwProfiler.vbwExecuteLine 753
     A2 = Asc(Mid$(Strg$, i, 1))
vbwProfiler.vbwExecuteLine 754
     A3 = A1 Xor A2
vbwProfiler.vbwExecuteLine 755
     b$ = Hex$(A3)
vbwProfiler.vbwExecuteLine 756
     If Len(b$) < 2 Then
vbwProfiler.vbwExecuteLine 757
          b$ = "0" + b$
     End If
vbwProfiler.vbwExecuteLine 758 'B
vbwProfiler.vbwExecuteLine 759
     s$ = s$ + b$
vbwProfiler.vbwExecuteLine 760
   Next
vbwProfiler.vbwExecuteLine 761
   EncryptINI$ = s$
vbwProfiler.vbwProcOut 23
vbwProfiler.vbwExecuteLine 762
End Function

Function DecryptINI$(ByVal Strg$, ByVal Password$)
vbwProfiler.vbwProcIn 24
   Dim b$, s$, i As Integer, j As Integer
   Dim A1 As Integer, A2 As Integer, A3 As Integer, P$
vbwProfiler.vbwExecuteLine 763
   j = 1
vbwProfiler.vbwExecuteLine 764
   For i = 1 To Len(Password$)
vbwProfiler.vbwExecuteLine 765
     P$ = P$ & Asc(Mid$(Password$, i, 1))
vbwProfiler.vbwExecuteLine 766
   Next

vbwProfiler.vbwExecuteLine 767
   For i = 1 To Len(Strg$) Step 2
vbwProfiler.vbwExecuteLine 768
     A1 = Asc(Mid$(P$, j, 1))
vbwProfiler.vbwExecuteLine 769
     j = j + 1
vbwProfiler.vbwExecuteLine 770
     If j > Len(P$) Then
vbwProfiler.vbwExecuteLine 771
          j = 1
     End If
vbwProfiler.vbwExecuteLine 772 'B
vbwProfiler.vbwExecuteLine 773
     b$ = Mid$(Strg$, i, 2)
vbwProfiler.vbwExecuteLine 774
     A3 = Val("&H" + b$)
vbwProfiler.vbwExecuteLine 775
     A2 = A1 Xor A3
vbwProfiler.vbwExecuteLine 776
     s$ = s$ + Chr$(A2)
vbwProfiler.vbwExecuteLine 777
   Next
vbwProfiler.vbwExecuteLine 778
   DecryptINI$ = s$
vbwProfiler.vbwProcOut 24
vbwProfiler.vbwExecuteLine 779
End Function


Sub ReadINI()
vbwProfiler.vbwProcIn 25
    Dim oFile As FileSystemObject
    Dim oINI As cINI
    ' fill our settings with default values first in case some of the INI values are missing
vbwProfiler.vbwExecuteLine 780
    With Settings
vbwProfiler.vbwExecuteLine 781
        .InitialDir = GVDPath
vbwProfiler.vbwExecuteLine 782
        .DesktopX = Screen.Width
vbwProfiler.vbwExecuteLine 783
        .DesktopY = Screen.Height
vbwProfiler.vbwExecuteLine 784
        .windowstate = vbNormal
vbwProfiler.vbwExecuteLine 785
        .FormTop = 0
vbwProfiler.vbwExecuteLine 786
        .FormLeft = 0
vbwProfiler.vbwExecuteLine 787
        .FormHeight = 600
vbwProfiler.vbwExecuteLine 788
        .FormWidth = 800
vbwProfiler.vbwExecuteLine 789
        .Splitter1 = 200
vbwProfiler.vbwExecuteLine 790
        .Splitter2 = 340
vbwProfiler.vbwExecuteLine 791
        .HSplitter = 340
vbwProfiler.vbwExecuteLine 792
        .bUseSurfaceAreaTable = True
vbwProfiler.vbwExecuteLine 793
        .bUseDefaultTextViewer = True
vbwProfiler.vbwExecuteLine 794
        .bUseDefaultWebBrowser = True
vbwProfiler.vbwExecuteLine 795
        .AuthorName = ""
vbwProfiler.vbwExecuteLine 796
        .Copyright = ""
vbwProfiler.vbwExecuteLine 797
        .email = ""
vbwProfiler.vbwExecuteLine 798
        .url = ""
vbwProfiler.vbwExecuteLine 799
        .Header = ""
vbwProfiler.vbwExecuteLine 800
        .Footer = ""
vbwProfiler.vbwExecuteLine 801
        .PublishEmailAddress = "veh@makosoft.com" 'todo: 02/16/02 Is this ok to have hardcoded?
vbwProfiler.vbwExecuteLine 802
        .DecimalPlaces = 2
vbwProfiler.vbwExecuteLine 803
        .FormatString = "standard"
        '.bQuickStart = False '02/16/02 MPJ (OBSOLETE)
        '.bSoundOff = False   '02/16/02 MPJ (obsolete)
vbwProfiler.vbwExecuteLine 804
        .bAssociateExt = False
vbwProfiler.vbwExecuteLine 805
        .TextExportPath = GVDPath
vbwProfiler.vbwExecuteLine 806
        .HTMLExportPath = GVDPath
vbwProfiler.vbwExecuteLine 807
        .VehiclesOpenPath = GVDPath
vbwProfiler.vbwExecuteLine 808
        .VehiclesSavePath = GVDPath
vbwProfiler.vbwExecuteLine 809
    End With

vbwProfiler.vbwExecuteLine 810
    Set oFile = New FileSystemObject
vbwProfiler.vbwExecuteLine 811
    Set oINI = New cINI

    ' Make sure the INI file exists
vbwProfiler.vbwExecuteLine 812
    If oFile.FileExists(GVDPath & "\" & GVDINIFile) Then

vbwProfiler.vbwExecuteLine 813
        oINI.FileName = GVDPath & "\" & GVDINIFile

        'read in the settings and store them into our Settings UDT
vbwProfiler.vbwExecuteLine 814
        With Settings

vbwProfiler.vbwExecuteLine 815
            .TextExportPath = oINI.ReadString("Paths", "TextExportPath")
vbwProfiler.vbwExecuteLine 816
            .HTMLExportPath = oINI.ReadString("Paths", "HTMLExportPath")
vbwProfiler.vbwExecuteLine 817
            .VehiclesOpenPath = oINI.ReadString("Paths", "VehiclesOpenPath")
vbwProfiler.vbwExecuteLine 818
            .VehiclesSavePath = oINI.ReadString("Paths", "VehiclesSavePath")

vbwProfiler.vbwExecuteLine 819
            .Recent1 = oINI.ReadString("Recent", "Recent1")
vbwProfiler.vbwExecuteLine 820
            .Recent2 = oINI.ReadString("Recent", "Recent2")
vbwProfiler.vbwExecuteLine 821
            .Recent3 = oINI.ReadString("Recent", "Recent3")
vbwProfiler.vbwExecuteLine 822
            .Recent4 = oINI.ReadString("Recent", "Recent4")
vbwProfiler.vbwExecuteLine 823
            .Recent5 = oINI.ReadString("Recent", "Recent5")

vbwProfiler.vbwExecuteLine 824
            .HTMLBrowserPath = oINI.ReadString("Viewers", "HTMLBrowserPath")
vbwProfiler.vbwExecuteLine 825
            .TextViewerPath = oINI.ReadString("Viewers", "TextViewerPath")

vbwProfiler.vbwExecuteLine 826
            .DesktopX = oINI.ReadInteger("Display", "DesktopX")
vbwProfiler.vbwExecuteLine 827
            .DesktopY = oINI.ReadInteger("Display", "DesktopY")
vbwProfiler.vbwExecuteLine 828
            .windowstate = oINI.ReadInteger("Display", "State")
vbwProfiler.vbwExecuteLine 829
            .FormTop = oINI.ReadInteger("Display", "Top")
vbwProfiler.vbwExecuteLine 830
            .FormLeft = oINI.ReadInteger("Display", "Left")
vbwProfiler.vbwExecuteLine 831
            .FormWidth = oINI.ReadInteger("Display", "Width")
vbwProfiler.vbwExecuteLine 832
            .FormHeight = oINI.ReadInteger("Display", "Height")
vbwProfiler.vbwExecuteLine 833
            .Splitter1 = oINI.ReadInteger("Display", "Splitter1")
vbwProfiler.vbwExecuteLine 834
            If .Splitter1 <= 0 Then
vbwProfiler.vbwExecuteLine 835
                 .Splitter1 = 4220
            End If
vbwProfiler.vbwExecuteLine 836 'B

vbwProfiler.vbwExecuteLine 837
            .Splitter2 = oINI.ReadInteger("Display", "Splitter2")
vbwProfiler.vbwExecuteLine 838
            If .Splitter2 <= 0 Then
vbwProfiler.vbwExecuteLine 839
                 .Splitter2 = 6720
            End If
vbwProfiler.vbwExecuteLine 840 'B

vbwProfiler.vbwExecuteLine 841
            .HSplitter = oINI.ReadInteger("Display", "HSplitter")
vbwProfiler.vbwExecuteLine 842
            If .HSplitter <= 0 Then
vbwProfiler.vbwExecuteLine 843
                 .HSplitter = 7000
            End If
vbwProfiler.vbwExecuteLine 844 'B

vbwProfiler.vbwExecuteLine 845
            .bUseSurfaceAreaTable = oINI.ReadInteger("Config", "UseSurfaceAreaTable")
vbwProfiler.vbwExecuteLine 846
            .bUseDefaultTextViewer = oINI.ReadInteger("Config", "UseDefaultTextViewer")
vbwProfiler.vbwExecuteLine 847
            .bUseDefaultWebBrowser = oINI.ReadInteger("Config", "UseDefaultWebBrowser")
vbwProfiler.vbwExecuteLine 848
            .PublishEmailAddress = oINI.ReadString("Config", "PublishEmail")
vbwProfiler.vbwExecuteLine 849
            .DecimalPlaces = oINI.ReadInteger("Config", "DecimalPlaces")
vbwProfiler.vbwExecuteLine 850
            .FormatString = oINI.ReadString("Config", "FormatCode")
            '.bQuickStart = oINI.ReadInteger("Config", "QuickStart") '02/16/02 MPJ (obsolete)
            '.bSoundOff = oINI.ReadInteger("Config", "DisableSound") '02/16/02 MPJ (obsolete)
vbwProfiler.vbwExecuteLine 851
            .bAssociateExt = oINI.ReadInteger("Config", "AssociateExt")

vbwProfiler.vbwExecuteLine 852
            .AuthorName = oINI.ReadString("Author", "Name")
vbwProfiler.vbwExecuteLine 853
            .email = oINI.ReadString("Author", "Email")
vbwProfiler.vbwExecuteLine 854
            .url = oINI.ReadString("Author", "URL")
vbwProfiler.vbwExecuteLine 855
            .Copyright = oINI.ReadString("Author", "Copyright")
vbwProfiler.vbwExecuteLine 856
            .Header = oINI.ReadString("Author", "Header")
vbwProfiler.vbwExecuteLine 857
            .Footer = oINI.ReadString("Author", "Footer")
vbwProfiler.vbwExecuteLine 858
        End With
    End If
vbwProfiler.vbwExecuteLine 859 'B


vbwProfiler.vbwProcOut 25
vbwProfiler.vbwExecuteLine 860
End Sub

Public Sub ReadLicenseFile()
vbwProfiler.vbwProcIn 26
    Dim oFile As FileSystemObject
    Dim iFree As Long

vbwProfiler.vbwExecuteLine 861
    On Error GoTo errorhandler

vbwProfiler.vbwExecuteLine 862
    Set oFile = New FileSystemObject

vbwProfiler.vbwExecuteLine 863
    iFree = FreeFile

vbwProfiler.vbwExecuteLine 864
    If oFile.FileExists(GVDPath & "\" & GVDLicenseFile) Then   ' debug  this line needs to change because its saving the INI in the wrong place
        'retreive the data from the file and store it into the Settings udt
vbwProfiler.vbwExecuteLine 865
        Open GVDPath + "\" + GVDLicenseFile For Binary As #iFree
vbwProfiler.vbwExecuteLine 866
        Get #iFree, 1, RegInfo
vbwProfiler.vbwExecuteLine 867
        Close #iFree

vbwProfiler.vbwExecuteLine 868
        With RegInfo
vbwProfiler.vbwExecuteLine 869
            gsRegName = .RegName
vbwProfiler.vbwExecuteLine 870
            gsRegID = .RegID
vbwProfiler.vbwExecuteLine 871
            gsRegNum = .RegNum
vbwProfiler.vbwExecuteLine 872
        End With
    End If
vbwProfiler.vbwExecuteLine 873 'B

vbwProfiler.vbwExecuteLine 874
    If IsEmpty(gsRegName) Then
vbwProfiler.vbwExecuteLine 875
        ReDim gsRegName(1)
    End If
vbwProfiler.vbwExecuteLine 876 'B
vbwProfiler.vbwExecuteLine 877
    If IsEmpty(gsRegNum) Then
vbwProfiler.vbwExecuteLine 878
        ReDim gsRegNum(1)
    End If
vbwProfiler.vbwExecuteLine 879 'B
vbwProfiler.vbwProcOut 26
vbwProfiler.vbwExecuteLine 880
    Exit Sub

errorhandler:
    ' must make sure our byte arrays are filled
vbwProfiler.vbwExecuteLine 881
    ReDim gsRegName(1)
vbwProfiler.vbwExecuteLine 882
    ReDim gsRegNum(1)

    'first close the file
vbwProfiler.vbwExecuteLine 883
    Close #iFree
    ' now delete it
vbwProfiler.vbwExecuteLine 884
    If oFile.FileExists(GVDPath & "\" & GVDLicenseFile) Then

vbwProfiler.vbwExecuteLine 885
        oFile.DeleteFile (GVDPath & "\" & GVDLicenseFile)
    End If
vbwProfiler.vbwExecuteLine 886 'B

vbwProfiler.vbwExecuteLine 887
    DoEvents

vbwProfiler.vbwProcOut 26
vbwProfiler.vbwExecuteLine 888
End Sub

Public Sub WriteLicenseFile()
vbwProfiler.vbwProcIn 27

    Dim iFree As Long

vbwProfiler.vbwExecuteLine 889
    iFree = FreeFile

    'first delete the file if it already exists
vbwProfiler.vbwExecuteLine 890
    Open GVDPath + "\" + GVDLicenseFile For Random As #iFree
vbwProfiler.vbwExecuteLine 891
    Close #iFree

vbwProfiler.vbwExecuteLine 892
    iFree = FreeFile

    ' open the License for binary write
vbwProfiler.vbwExecuteLine 893
    Open GVDPath + "\" + GVDLicenseFile For Binary As #iFree

    ' update the relevant settings before we save it
vbwProfiler.vbwExecuteLine 894
    With RegInfo
vbwProfiler.vbwExecuteLine 895
        .RegID = gsRegID
vbwProfiler.vbwExecuteLine 896
        .RegName = gsRegName
vbwProfiler.vbwExecuteLine 897
        .RegNum = gsRegNum
vbwProfiler.vbwExecuteLine 898
    End With

    ' save the Settings data and close the file
vbwProfiler.vbwExecuteLine 899
    Put #iFree, , RegInfo
vbwProfiler.vbwExecuteLine 900
    Close #iFree

vbwProfiler.vbwProcOut 27
vbwProfiler.vbwExecuteLine 901
End Sub

Sub WriteINI()
vbwProfiler.vbwProcIn 28
vbwProfiler.vbwExecuteLine 902
    On Error Resume Next
    Dim oINI As cINI
vbwProfiler.vbwExecuteLine 903
    Set oINI = New cINI

vbwProfiler.vbwExecuteLine 904
    oINI.FileName = GVDPath & "\" & GVDINIFile

vbwProfiler.vbwExecuteLine 905
    With Settings
vbwProfiler.vbwExecuteLine 906
        Call oINI.WriteInteger("Display", "DesktopX", Screen.Width)
vbwProfiler.vbwExecuteLine 907
        Call oINI.WriteInteger("Display", "DesktopY", Screen.Height)
vbwProfiler.vbwExecuteLine 908
        Call oINI.WriteInteger("Display", "State", frmDesigner.windowstate)

        'JAW 2000.05.22
        'Splitter positions were not being saved when window was maximized.
        'If (frmDesigner.windowstate <> vbMaximized) And (frmDesigner.windowstate <> vbMinimized) Then
vbwProfiler.vbwExecuteLine 909
        If (frmDesigner.windowstate <> vbMinimized) Then
vbwProfiler.vbwExecuteLine 910
            Call oINI.WriteInteger("Display", "Top", Settings.FormTop)
vbwProfiler.vbwExecuteLine 911
            Call oINI.WriteInteger("Display", "Left", Settings.FormLeft)
vbwProfiler.vbwExecuteLine 912
            Call oINI.WriteInteger("Display", "Width", Settings.FormWidth)
vbwProfiler.vbwExecuteLine 913
            Call oINI.WriteInteger("Display", "Height", Settings.FormHeight)
vbwProfiler.vbwExecuteLine 914
            Call oINI.WriteInteger("Display", "Splitter1", Settings.Splitter1)
           ' Call oINI.WriteInteger("Display", "Splitter2", frmDesigner.ListView1.Left + frmDesigner.ListView1.Width)
vbwProfiler.vbwExecuteLine 915
            Call oINI.WriteInteger("Display", "HSplitter", Settings.HSplitter)
        End If
vbwProfiler.vbwExecuteLine 916 'B

vbwProfiler.vbwExecuteLine 917
        Call oINI.WriteInteger("Config", "UseSurfaceAreaTable", Settings.bUseSurfaceAreaTable)
vbwProfiler.vbwExecuteLine 918
        Call oINI.WriteInteger("Config", "UseDefaultTextViewer", Settings.bUseDefaultTextViewer)
vbwProfiler.vbwExecuteLine 919
        Call oINI.WriteInteger("Config", "UseDefaultWebBrowser", Settings.bUseDefaultWebBrowser)
vbwProfiler.vbwExecuteLine 920
        Call oINI.WriteString("Config", "PublishEmail", .PublishEmailAddress)
vbwProfiler.vbwExecuteLine 921
        Call oINI.WriteInteger("Config", "DecimalPlaces", .DecimalPlaces)
vbwProfiler.vbwExecuteLine 922
        Call oINI.WriteString("Config", "FormatCode", "standard")
       ' Call oINI.WriteInteger("Config", "QuickStart", Settings.bQuickStart) '02/16/02 MPJ (obsolete)
       ' Call oINI.WriteInteger("Config", "DisableSound", Settings.bSoundOff) '02/16/02 MPJ (obsolete)
vbwProfiler.vbwExecuteLine 923
        Call oINI.WriteInteger("Config", "AssociateExt", Settings.bAssociateExt)

vbwProfiler.vbwExecuteLine 924
        Call oINI.WriteString("Author", "Name", Settings.AuthorName)
vbwProfiler.vbwExecuteLine 925
        Call oINI.WriteString("Author", "Email", Settings.email)
vbwProfiler.vbwExecuteLine 926
        Call oINI.WriteString("Author", "URL", Settings.url)
vbwProfiler.vbwExecuteLine 927
        Call oINI.WriteString("Author", "Copyright", Settings.Copyright)
vbwProfiler.vbwExecuteLine 928
        Call oINI.WriteString("Author", "Header", Settings.Header)
vbwProfiler.vbwExecuteLine 929
        Call oINI.WriteString("Author", "Footer", Settings.Footer)

vbwProfiler.vbwExecuteLine 930
        Call oINI.WriteString("Paths", "App", GVDPath)
vbwProfiler.vbwExecuteLine 931
        Call oINI.WriteString("Paths", "TextExportPath", .TextExportPath)
vbwProfiler.vbwExecuteLine 932
        Call oINI.WriteString("Paths", "HTMLExportPath", .HTMLExportPath)
vbwProfiler.vbwExecuteLine 933
        Call oINI.WriteString("Paths", "VehiclesOpenPath", .VehiclesOpenPath)
vbwProfiler.vbwExecuteLine 934
        Call oINI.WriteString("Paths", "VehiclesSavePath", .VehiclesSavePath)

vbwProfiler.vbwExecuteLine 935
        Call oINI.WriteString("Viewers", "HTMLBrowserPath", .HTMLBrowserPath)
vbwProfiler.vbwExecuteLine 936
        Call oINI.WriteString("Viewers", "TextViewerPath", .TextViewerPath)

vbwProfiler.vbwExecuteLine 937
        Call oINI.WriteString("Recent", "Recent1", frmDesigner.mnuRecent(1).Caption)
vbwProfiler.vbwExecuteLine 938
        Call oINI.WriteString("Recent", "Recent2", frmDesigner.mnuRecent(2).Caption)
vbwProfiler.vbwExecuteLine 939
        Call oINI.WriteString("Recent", "Recent3", frmDesigner.mnuRecent(3).Caption)
vbwProfiler.vbwExecuteLine 940
        Call oINI.WriteString("Recent", "Recent4", frmDesigner.mnuRecent(4).Caption)
vbwProfiler.vbwExecuteLine 941
        Call oINI.WriteString("Recent", "Recent5", frmDesigner.mnuRecent(5).Caption)
vbwProfiler.vbwExecuteLine 942
    End With
vbwProfiler.vbwProcOut 28
vbwProfiler.vbwExecuteLine 943
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


