Attribute VB_Name = "modGlobals"
Option Explicit

'======================================================
' Constants
'======================================================
'File    :  modGlobals.bas
'Created :  July 10th 2002
'Author  :  Mike Joseph
'Notes   :  Finally seperating out most of the global variables into a single file.
'           There are still very few left in modules like modHeaps but those particular globals are used sparingly in the clsComplst only


'======================================================
' Constants
'======================================================
'#Const DEBUG_MODE = True
Public Const Password = "psswdGVD1999CopyrightMichaelJoseph"

Public Const CHECKLIST_STATE_RESTORE = "RESTORE"


Public Const SW_SHOW = 5        ' Displays Window in its current size and position
Public Const SW_SHOWNORMAL = 1
Public Const SE_ERR_FNF = 2&
Public Const SE_ERR_PNF = 3&
Public Const SE_ERR_ACCESSDENIED = 5&
Public Const SE_ERR_OOM = 8&
Public Const SE_ERR_DLLNOTFOUND = 32&
Public Const SE_ERR_SHARE = 26&
Public Const SE_ERR_ASSOCINCOMPLETE = 27&
Public Const SE_ERR_DDETIMEOUT = 28&
Public Const SE_ERR_DDEFAIL = 29&
Public Const SE_ERR_DDEBUSY = 30&
Public Const SE_ERR_NOASSOC = 31&
Public Const ERROR_BAD_FORMAT = 11&


'======================================================
' Global vars
'======================================================
Public Settings As udtSettings 'todo: change this to an object which uses cIPersist to handle save/load?

'Public p_lngDataType As Integer '07/10/02 MPJ <-- Christ, this has been here forever.  It was named Datatype but ive just now changed it to p_lngDatatype... so far no consequences observed TODO: Safe to delete?
'Public p_lngImageIndex As Integer  ' index of imagelist icon
'Public p_nIndex As Integer  ' Holds the index of a Node
Public p_bChangedFlag As Boolean ' JAW 2000.05.07 tracks whether .veh has been changed since last loaded or saved

Public GVDVehiclesPath As String
Public GVDPath As String

'======================================================
' Active TreeVehicle Node Tracking Type and Global
'======================================================
'todo: is this being used???  Its not even declared public???
Private Type trkNode
    Key As String
    Parent As String
    Datatype As Long ' also functions as node type for non component nodes like Performance/WeaponLinks/Profiles
    ParentDataType As Long
    CustomDescription As String
    'TODO: I might even be able to have a member IsDeleteAble as Boolean and perhaps even pDeleteFunction
    ' for storing a pointer to the function which will delete the componet if its deletable.  This way
    ' i can simplify the code to where only the SetActiveNode has to decipher which type of node it is
    ' and what is/is not allowed to be done with that type of node
End Type

Public p_ActiveNode As trkNode

Public Type udtComponent
    Classname As String
    ComponentPath As String
    DefPath As String
    GUID As String
    'lPtr As Long
    IconPath As String
    Text As String
End Type

'======================================================
Public gsMajor As String
Public gsMinor As String
Public gsRevision As String
Public gsRegID As Long
Public gsRegNum() As Byte
Public gsRegName() As Byte
Public p_sGUID As String * 39


Private Type udtSettings
     PublishEmailAddress As String
     HTMLBrowserPath As String
     TextViewerPath As String
     ProgramVersion As String
     SerialNumber As String
     InitialDir As String
     DesktopX As Long
     DesktopY As Long
     windowstate As Long
     FormTop As Long
     FormLeft As Long
     FormHeight As Long
     FormWidth As Long
     Splitter1 As Long
     Splitter2 As Long
     HSplitter As Long
     bUseSurfaceAreaTable As Boolean
     bUseDefaultWebBrowser As Boolean
     bUseDefaultTextViewer As Boolean
     'bSoundOff As Boolean       'MPJ 02/16/02 Obsolete
     'bQuickStart As Boolean     'MPJ 02/16/02 Obsolete
     bAssociateExt As Boolean
     AuthorName As String
     Copyright As String
     email As String
     url As String
     Header As String
     Footer As String
     DecimalPlaces As Integer
     FormatString As String
     TextExportPath As String
     HTMLExportPath As String
     VehiclesOpenPath As String
     VehiclesSavePath As String
     Recent1 As String
     Recent2 As String
     Recent3 As String
     Recent4 As String
     Recent5 As String
End Type


