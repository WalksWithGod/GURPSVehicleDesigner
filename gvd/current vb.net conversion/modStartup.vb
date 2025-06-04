Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module modStartup
	
	' DECLARES
	Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
	Private Declare Function timeGetTime Lib "winmm.dll" () As Integer
	
	Private Const INSTANCE_RUNNING As String = "An instance of GURPS Vehicle Designer 2.0 is already running."
	
#If DEBUG_MODE Then
	'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression DEBUG_MODE did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
	Private Const SLEEP_TIME = 0 ' milliseconds
#Else
	Private Const SLEEP_TIME As Short = 2500
#End If
	
	Sub UnloadAllForms(ByVal sFormName As String)
		
		Dim Form As System.Windows.Forms.Form
		For	Each Form In My.Application.OpenForms
			If Form.Name <> sFormName Then
				Form.Close()
				'UPGRADE_NOTE: Object Form may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
				Form = Nothing
			End If
		Next Form
		
		Reset() ' close any open files
	End Sub
	
	'todo: this eventually needs to accept an index or handle to a specific veh in our vehicle collection
	'todo: actually, im looking for a way where this update doesnt get called from the client since eventually this should be
	' done server side.  Client has no right to initiate an update.
	Sub UpdateVehicle(ByVal lptrVehicle As Integer, ByVal lptrNode As Integer)
		Dim Stats As Object
		Dim frmDesigner As Object
		Dim i As Integer
		Dim keyarray() As String
		Dim element As Object
		Dim oVehicle As cVehicle
		
		
		' add an hour glass pointer  '- 06/03/2000 MPJ
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MousePointer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.MousePointer = System.Windows.Forms.Cursors.WaitCursor
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.Enabled. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.Enabled = False
		' This is the only subroutine i need to use to update stats
		' update the Vehicle's Total  Stats
		
		On Error Resume Next
		'UPGRADE_WARNING: Couldn't resolve default property of object oVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(oVehicle, lptrVehicle, 4)
		
		' todo:, do i need to do a search for the "stats" component in this vehicle?
		' suddenly ive forgotten why we cant just access oVehicle.Stats.xxxx
		' Actually, im thinking that cVehicle now hosts these basic overall stats itself
		' cStats has been moved to a new shared project (DLL) and is used for actually
		' "calculating" stats.  But for vehicle. these overall stats are inherent
		
		'---------------
		' todo: since we are only updating ONLY the components which have been changed
		'       and those which are effected by the change (e.g. a parent after its child has been changed)
		'       we need a different function to handle that.  For now lets do it here
		Dim oStats As Stats.cStats
		Dim oNode As _cINode
		Dim oComponent As _cIComponent
		
		'UPGRADE_WARNING: Couldn't resolve default property of object oNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(oNode, lptrNode, 4)
		
		oStats = New cStats
		'UPGRADE_WARNING: Couldn't resolve default property of object oComponent.calcStats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oComponent.calcStats(oStats)
		
		
		'UPGRADE_NOTE: Object oStats may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oStats = Nothing
		'UPGRADE_NOTE: Object oComponent may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oComponent = Nothing
		'UPGRADE_WARNING: Couldn't resolve default property of object oNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(oNode, 0, 4)
		'------------------
		
		' display the totals in the stats pane
		'    With oVehicle
		'        frmDesigner.lstStats.Clear
		'        frmDesigner.lstStats.Columns = 3
		'
		'        'todo: i should provide user with font dialog
		'        'frmDesigner.lstStats.Font = frmDesigner.PLC1.Font
		'        'frmDesigner.lstStats.FontSize = frmDesigner.PLC1.Font.size
		'
		'        frmDesigner.lstStats.AddItem "Price: " & Format(.TotalPrice, "standard") 'todo: we should just refer to this as .Cost
		'        frmDesigner.lstStats.AddItem "Health: " & .StructuralHealth & " HT"
		'        frmDesigner.lstStats.AddItem "SizeMod: " & .SizeModifier
		'        frmDesigner.lstStats.AddItem "Volume: " & Format(.TotalVolume, "standard") & " cu ft" 'todo: should just refer to as .Volume
		'        frmDesigner.lstStats.AddItem "Area: " & Format(.totalsurfacearea, "standard") & " sq ft" 'todo: .surfacearea
		'        frmDesigner.lstStats.AddItem "Empty Wt: " & Format(.EmptyWeight, "standard") & " lbs" 'todo: weight
		'        frmDesigner.lstStats.AddItem "Empty Mass: " & Format(.EmptyWeight / 2000, "standard") & " tons"
		'        frmDesigner.lstStats.AddItem "Loaded Wt: " & Format(.LoadedWeight, "standard") & " lbs"
		'        frmDesigner.lstStats.AddItem "Loaded Mass: " & Format(.LoadedMass, "standard") & " tons"
		'        frmDesigner.lstStats.AddItem "+Hardpoint Wt: " & Format(.HLoadedWeight, "standard") & " lbs"
		'        frmDesigner.lstStats.AddItem "+Hardpoint Mass: " & Format(.HLoadedMass, "standard") & " tons"
		'        frmDesigner.lstStats.AddItem "Submerged Wt: " & Format(.SubmergedWeight, "standard") & " lbs"
		'        frmDesigner.lstStats.AddItem "Submerged Mass: " & Format(.SubmergedMass, "standard") & " tons"
		'        frmDesigner.lstStats.AddItem "Continuous Power Output: " & Format(.TotalContinuousPower, "standard") & " kW"
		'        frmDesigner.lstStats.AddItem "Continuous Power Consumption: " & Format(.TotalContinuousPowerConsumption, "standard") & " kW"
		'         frmDesigner.lstStats.AddItem "Stored Power Output: " & Format(.TotalstoredPower, "standard") & " kWs"
		'        frmDesigner.lstStats.AddItem "Stored Power Consumption: " & Format(.TotalStoredPowerConsumption, "standard") & " kWs"
		'
		'    End With
		
		'UPGRADE_WARNING: Couldn't resolve default property of object oVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(oVehicle, 0, 4)
		
		Properties_Show(p_ActiveNode.Key, p_ActiveNode.Datatype)
		DisplayPrintOutput()
		
		
		' reset mouse
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.Enabled. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.Enabled = True
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.MousePointer. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
		frmDesigner.MousePointer = vbNormal
	End Sub
	
	
	
	'UPGRADE_NOTE: Main was upgraded to Main_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Sub Main_Renamed()
		Dim frmDesigner As Object
		Dim frmSplash As Object
		
		Dim i As Integer ' debug delay counter to keep the splash screen visible longer
		'UPGRADE_ISSUE: FileSystemObject object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim oFile As FileSystemObject
		Dim sFile As String
		Dim lngTime As Integer
		
		'abort if GVD is already running
		'UPGRADE_ISSUE: App property App.PrevInstance was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="076C26E5-B7A9-4E77-B69C-B4448DF39E58"'
		If App.PrevInstance Then
			If VB.Command() <> "" Then
				'pass the vehicle to the other instance of the application and give the user
				'the opportunity to open it
				
			Else
				MsgBox(INSTANCE_RUNNING)
				End
			End If
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object frmSplash.Show. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmSplash.Show()
		lngTime = timeGetTime
		System.Windows.Forms.Application.DoEvents()
		''''''''''''''''''''''''''''''
		'02/16/02 MPJ Removing fancy splash screen, hence no longer need to refresh
		' or doevents to make sure everything displays properly
		'frmSplash.Refresh
		'DoEvents
		'''''''''''''''''''''''''''''''
		
		GVDPath = My.Application.Info.DirectoryPath ' get the apps directory.
		ChDir(GVDPath) 'make sure we are in the right directory
		
		' read the ini file (This should use a REAL ini file)
		ReadINI()
		ReadLicenseFile()
		
		''''''''''''''''''''''''''''''''''''''''''''
		'02/16/02 MPJ - WAV Sounds were just annoying.  Killing them altogether (plus
		' with the simpler splash screen, it no longer makes sense anyway
		'If Not Settings.bSoundOff Then
		'    SoundName$ = App.Path & "\intro.wav"
		'    wFlags% = SND_ASYNC Or SND_NODEFAULT 'if it cant find / play the soundfile.. it will simply resume w/out sound
		'    x% = sndPlaySound(SoundName$, wFlags%)
		'End If
		''''''''''''''''''''''''''''''''''''''''''''
		
		' load the main form
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		Load(frmDesigner)
		
		sFile = VB.Command() ' our command line argument
		If sFile <> "" Then
			oFile = New FileSystemObject
			'UPGRADE_WARNING: Couldn't resolve default property of object oFile.FileExists. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If oFile.FileExists(sFile) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.LoadVehicle. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If frmDesigner.LoadVehicle(sFile) Then
					setGUID()
				End If
			End If
			'UPGRADE_NOTE: Object oFile may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oFile = Nothing
		Else
			'load a new vehicle
			'frmDesigner.LoadNewVehicle  'todo: Leave this commented or no?  See how users feel....
		End If
		
		
		''''''''''''''''''''''''''''''''''''''''
		'02/16/02 MPJ No more fancy splashscreen
		' If Not Settings.bQuickStart Then
		'     Pause 1.5
		'     FlashBmps
		'     Pause 1.5
		' End If
		' Doevents
		''''''''''''''''''''''''''''''''''''''
		' make sure the splash screen is visible for a minimum of SLEEP_TIME so that
		' the SJGames copyright and trademark info is visible long enuf
		lngTime = timeGetTime - lngTime
		If lngTime < SLEEP_TIME Then
			Sleep(SLEEP_TIME - lngTime)
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.Show. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.Show()
		'UPGRADE_ISSUE: Unload frmSplash was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="875EBAD7-D704-4539-9969-BC7DBDAA62A2"'
		Unload(frmSplash)
		'UPGRADE_NOTE: Object frmSplash may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		frmSplash = Nothing 'todo: this is totally not needed im sure.  Why is it still here?  Think its leftover from some old changes
		
	End Sub
End Module