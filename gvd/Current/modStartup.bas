Attribute VB_Name = "modStartup"
Option Explicit

' DECLARES
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Const INSTANCE_RUNNING = "An instance of GURPS Vehicle Designer 2.0 is already running."

#If DEBUG_MODE Then
    Private Const SLEEP_TIME = 0 ' milliseconds
#Else
    Private Const SLEEP_TIME = 2500
#End If

Sub UnloadAllForms(ByVal sFormName As String)

    Dim Form As Form
   For Each Form In Forms
      If Form.Name <> sFormName Then
         Unload Form
         Set Form = Nothing
      End If
   Next Form
   
   Reset ' close any open files
End Sub

'todo: this eventually needs to accept an index or handle to a specific veh in our vehicle collection
'todo: actually, im looking for a way where this update doesnt get called from the client since eventually this should be
' done server side.  Client has no right to initiate an update.
Sub UpdateVehicle(ByVal lptrVehicle As Long, ByVal lptrNode As Long)
    Dim i As Long
    Dim keyarray() As String
    Dim element As Object
    Dim oVehicle As cVehicle
    
    
    ' add an hour glass pointer  '- 06/03/2000 MPJ
    frmDesigner.MousePointer = vbHourglass
    frmDesigner.Enabled = False
    ' This is the only subroutine i need to use to update stats
    ' update the Vehicle's Total  Stats
    
    On Error Resume Next
    CopyMemory oVehicle, lptrVehicle, 4
    
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
    Dim oNode As cINode
    Dim oComponent As cIComponent
    
    CopyMemory oNode, lptrNode, 4
    
    Set oStats = New cStats
    oComponent.calcStats oStats
    
    
    Set oStats = Nothing
    Set oComponent = Nothing
    CopyMemory oNode, 0&, 4
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
    
    CopyMemory oVehicle, 0&, 4
    
    Properties_Show p_ActiveNode.Key, p_ActiveNode.Datatype
    DisplayPrintOutput

    
    ' reset mouse
    frmDesigner.Enabled = True
    frmDesigner.MousePointer = vbNormal
End Sub



Sub Main()

    Dim i As Long ' debug delay counter to keep the splash screen visible longer
    Dim oFile As FileSystemObject
    Dim sFile As String
    Dim lngTime As Long
    
    'abort if GVD is already running
    If App.PrevInstance Then
        If Command <> "" Then
            'pass the vehicle to the other instance of the application and give the user
            'the opportunity to open it
            
        Else
            MsgBox INSTANCE_RUNNING
            End
        End If
    End If

    frmSplash.Show
    lngTime = timeGetTime
    DoEvents
    ''''''''''''''''''''''''''''''
    '02/16/02 MPJ Removing fancy splash screen, hence no longer need to refresh
    ' or doevents to make sure everything displays properly
    'frmSplash.Refresh
    'DoEvents
    '''''''''''''''''''''''''''''''
    
    GVDPath = App.Path ' get the apps directory.
    ChDir GVDPath 'make sure we are in the right directory
    
    ' read the ini file (This should use a REAL ini file)
    ReadINI
    ReadLicenseFile
    
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
    Load frmDesigner
    
     sFile = Command ' our command line argument
    If sFile <> "" Then
        Set oFile = New FileSystemObject
        If oFile.FileExists(sFile) Then
            If frmDesigner.LoadVehicle(sFile) Then
                setGUID
            End If
        End If
        Set oFile = Nothing
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
        Sleep SLEEP_TIME - lngTime
    End If
    frmDesigner.Show
    Unload frmSplash
    Set frmSplash = Nothing  'todo: this is totally not needed im sure.  Why is it still here?  Think its leftover from some old changes
    
End Sub
