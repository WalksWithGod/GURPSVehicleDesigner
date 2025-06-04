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
vbwProfiler.vbwProcIn 244

    Dim Form As Form
vbwProfiler.vbwExecuteLine 4637
   For Each Form In Forms
vbwProfiler.vbwExecuteLine 4638
      If Form.Name <> sFormName Then
vbwProfiler.vbwExecuteLine 4639
         Unload Form
vbwProfiler.vbwExecuteLine 4640
         Set Form = Nothing
      End If
vbwProfiler.vbwExecuteLine 4641 'B
vbwProfiler.vbwExecuteLine 4642
   Next Form

vbwProfiler.vbwExecuteLine 4643
   Reset ' close any open files
vbwProfiler.vbwProcOut 244
vbwProfiler.vbwExecuteLine 4644
End Sub

'todo: this eventually needs to accept an index or handle to a specific veh in our vehicle collection
'todo: actually, im looking for a way where this update doesnt get called from the client since eventually this should be
' done server side.  Client has no right to initiate an update.
Sub UpdateVehicle(ByVal lptrVehicle As Long, ByVal lptrNode As Long)
vbwProfiler.vbwProcIn 245
    Dim i As Long
    Dim keyarray() As String
    Dim element As Object
    Dim oVehicle As cVehicle


    ' add an hour glass pointer  '- 06/03/2000 MPJ
vbwProfiler.vbwExecuteLine 4645
    frmDesigner.MousePointer = vbHourglass
vbwProfiler.vbwExecuteLine 4646
    frmDesigner.Enabled = False
    ' This is the only subroutine i need to use to update stats
    ' update the Vehicle's Total  Stats

vbwProfiler.vbwExecuteLine 4647
    On Error Resume Next
vbwProfiler.vbwExecuteLine 4648
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

vbwProfiler.vbwExecuteLine 4649
    CopyMemory oNode, lptrNode, 4

vbwProfiler.vbwExecuteLine 4650
    Set oStats = New cStats
vbwProfiler.vbwExecuteLine 4651
    oComponent.calcStats oStats


vbwProfiler.vbwExecuteLine 4652
    Set oStats = Nothing
vbwProfiler.vbwExecuteLine 4653
    Set oComponent = Nothing
vbwProfiler.vbwExecuteLine 4654
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

vbwProfiler.vbwExecuteLine 4655
    CopyMemory oVehicle, 0&, 4

vbwProfiler.vbwExecuteLine 4656
    Properties_Show p_ActiveNode.Key, p_ActiveNode.Datatype
vbwProfiler.vbwExecuteLine 4657
    DisplayPrintOutput


    ' reset mouse
vbwProfiler.vbwExecuteLine 4658
    frmDesigner.Enabled = True
vbwProfiler.vbwExecuteLine 4659
    frmDesigner.MousePointer = vbNormal
vbwProfiler.vbwProcOut 245
vbwProfiler.vbwExecuteLine 4660
End Sub



Sub Main()
    vbwInitializeProfiler ' Initialize VB Watch
vbwProfiler.vbwProcIn 246

    Dim i As Long ' debug delay counter to keep the splash screen visible longer
    Dim oFile As FileSystemObject
    Dim sFile As String
    Dim lngTime As Long

    'abort if GVD is already running
vbwProfiler.vbwExecuteLine 4661
    If App.PrevInstance Then
vbwProfiler.vbwExecuteLine 4662
        If Command <> "" Then
            'pass the vehicle to the other instance of the application and give the user
            'the opportunity to open it

        Else
vbwProfiler.vbwExecuteLine 4663 'B
vbwProfiler.vbwExecuteLine 4664
            MsgBox INSTANCE_RUNNING
vbwProfiler.vbwProcOut 246
'vbwLine 4665:            End
vbwProfiler.vbwExecuteLine 4665: Call vbwFinishSession: End
        End If
vbwProfiler.vbwExecuteLine 4666 'B
    End If
vbwProfiler.vbwExecuteLine 4667 'B

vbwProfiler.vbwExecuteLine 4668
    frmSplash.Show
vbwProfiler.vbwExecuteLine 4669
    lngTime = timeGetTime
vbwProfiler.vbwExecuteLine 4670
    DoEvents
    ''''''''''''''''''''''''''''''
    '02/16/02 MPJ Removing fancy splash screen, hence no longer need to refresh
    ' or doevents to make sure everything displays properly
    'frmSplash.Refresh
    'DoEvents
    '''''''''''''''''''''''''''''''

vbwProfiler.vbwExecuteLine 4671
    GVDPath = App.Path ' get the apps directory.
vbwProfiler.vbwExecuteLine 4672
    ChDir GVDPath 'make sure we are in the right directory

    ' read the ini file (This should use a REAL ini file)
vbwProfiler.vbwExecuteLine 4673
    ReadINI
vbwProfiler.vbwExecuteLine 4674
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
vbwProfiler.vbwExecuteLine 4675
    Load frmDesigner

vbwProfiler.vbwExecuteLine 4676
     sFile = Command ' our command line argument
vbwProfiler.vbwExecuteLine 4677
    If sFile <> "" Then
vbwProfiler.vbwExecuteLine 4678
        Set oFile = New FileSystemObject
vbwProfiler.vbwExecuteLine 4679
        If oFile.FileExists(sFile) Then
vbwProfiler.vbwExecuteLine 4680
            If frmDesigner.LoadVehicle(sFile) Then
vbwProfiler.vbwExecuteLine 4681
                setGUID
            End If
vbwProfiler.vbwExecuteLine 4682 'B
        End If
vbwProfiler.vbwExecuteLine 4683 'B
vbwProfiler.vbwExecuteLine 4684
        Set oFile = Nothing
    Else
vbwProfiler.vbwExecuteLine 4685 'B
        'load a new vehicle
        'frmDesigner.LoadNewVehicle  'todo: Leave this commented or no?  See how users feel....
    End If
vbwProfiler.vbwExecuteLine 4686 'B


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
vbwProfiler.vbwExecuteLine 4687
    lngTime = timeGetTime - lngTime
vbwProfiler.vbwExecuteLine 4688
    If lngTime < SLEEP_TIME Then
vbwProfiler.vbwExecuteLine 4689
        Sleep SLEEP_TIME - lngTime
    End If
vbwProfiler.vbwExecuteLine 4690 'B
vbwProfiler.vbwExecuteLine 4691
    frmDesigner.Show
vbwProfiler.vbwExecuteLine 4692
    Unload frmSplash
vbwProfiler.vbwExecuteLine 4693
    Set frmSplash = Nothing  'todo: this is totally not needed im sure.  Why is it still here?  Think its leftover from some old changes

vbwProfiler.vbwProcOut 246
vbwProfiler.vbwExecuteLine 4694
End Sub


