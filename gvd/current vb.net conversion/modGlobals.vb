Option Strict Off
Option Explicit On
Module modGlobals
	
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
	Public Const Password As String = "psswdGVD1999CopyrightMichaelJoseph"
	
	Public Const CHECKLIST_STATE_RESTORE As String = "RESTORE"
	
	
	Public Const SW_SHOW As Short = 5 ' Displays Window in its current size and position
	Public Const SW_SHOWNORMAL As Short = 1
	Public Const SE_ERR_FNF As Short = 2
	Public Const SE_ERR_PNF As Short = 3
	Public Const SE_ERR_ACCESSDENIED As Short = 5
	Public Const SE_ERR_OOM As Short = 8
	Public Const SE_ERR_DLLNOTFOUND As Short = 32
	Public Const SE_ERR_SHARE As Short = 26
	Public Const SE_ERR_ASSOCINCOMPLETE As Short = 27
	Public Const SE_ERR_DDETIMEOUT As Short = 28
	Public Const SE_ERR_DDEFAIL As Short = 29
	Public Const SE_ERR_DDEBUSY As Short = 30
	Public Const SE_ERR_NOASSOC As Short = 31
	Public Const ERROR_BAD_FORMAT As Short = 11
	
	
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
	Private Structure trkNode
		Dim Key As String
		Dim Parent As String
		Dim Datatype As Integer ' also functions as node type for non component nodes like Performance/WeaponLinks/Profiles
		Dim ParentDataType As Integer
		Dim CustomDescription As String
		'TODO: I might even be able to have a member IsDeleteAble as Boolean and perhaps even pDeleteFunction
		' for storing a pointer to the function which will delete the componet if its deletable.  This way
		' i can simplify the code to where only the SetActiveNode has to decipher which type of node it is
		' and what is/is not allowed to be done with that type of node
	End Structure
	
	Public p_ActiveNode As trkNode
	
	Public Structure udtComponent
		Dim Classname As String
		Dim ComponentPath As String
		Dim DefPath As String
		Dim GUID As String
		'lPtr As Long
		Dim IconPath As String
		Dim Text As String
	End Structure
	
	'======================================================
	Public gsMajor As String
	Public gsMinor As String
	Public gsRevision As String
	Public gsRegID As Integer
	Public gsRegNum() As Byte
	Public gsRegName() As Byte
	Public p_sGUID As New VB6.FixedLengthString(39)
	
	
	Private Structure udtSettings
		Dim PublishEmailAddress As String
		Dim HTMLBrowserPath As String
		Dim TextViewerPath As String
		Dim ProgramVersion As String
		Dim SerialNumber As String
		Dim InitialDir As String
		Dim DesktopX As Integer
		Dim DesktopY As Integer
		Dim windowstate As Integer
		Dim FormTop As Integer
		Dim FormLeft As Integer
		Dim FormHeight As Integer
		Dim FormWidth As Integer
		Dim Splitter1 As Integer
		Dim Splitter2 As Integer
		Dim HSplitter As Integer
		Dim bUseSurfaceAreaTable As Boolean
		Dim bUseDefaultWebBrowser As Boolean
		Dim bUseDefaultTextViewer As Boolean
		'bSoundOff As Boolean       'MPJ 02/16/02 Obsolete
		'bQuickStart As Boolean     'MPJ 02/16/02 Obsolete
		Dim bAssociateExt As Boolean
		Dim AuthorName As String
		Dim Copyright As String
		Dim email As String
		Dim url As String
		Dim Header As String
		Dim Footer As String
		Dim DecimalPlaces As Short
		Dim FormatString As String
		Dim TextExportPath As String
		Dim HTMLExportPath As String
		Dim VehiclesOpenPath As String
		Dim VehiclesSavePath As String
		Dim Recent1 As String
		Dim Recent2 As String
		Dim Recent3 As String
		Dim Recent4 As String
		Dim Recent5 As String
	End Structure
End Module