Attribute VB_Name = "ModMain"
Option Explicit
Public Datatype As Integer
Public myImage As Integer  ' index of imagelist icon
Public mySelectedImage As Integer  ' index of the selected image icon
Public TableName As String
Public fMainForm As frmDesigner
Public colVehicle As New Collection
Dim objNode As node
Public mnIndex As Integer  ' Holds the index of a Node
Public ArmorOption As Integer
Public gVehicleTL As Integer ' the vehicles tech level
Const Overall = 0
Const Location = 1
Const Basic = 2
Const Complex = 3
Const Skirt = 4
Const GunShield = 5
Const OpenFrame = 6

'//non file handling UDT's.  See the modFileHandling for the other UDTs
Public Type udtThruster
    ClassName As String
    Key As String
    Checked As Boolean
End Type

'==================================================
'==================================================
'==================================================
'==================================================
'==================================================
' Class declarations for Sub Assemblies
'==================================================
Public objBody As New clsBody
Public objWheel As New clsWheel
Public objSuperStructure As New clsSuperStructure
Public objWing As New clsWing
Public objRotar As New clsRotar
Public objHydrofoil As New clsHydrofoil
Public objHovercraft As New clsHovercraft
Public objSkid As New clsSkid
Public objTrack As New clsTrack
Public objLeg As New clsLeg
Public objOpenMount As New clsOpenMount
Public objMast As New clsMast
Public objPod As New clsPod
Public objTurret As New clsTurret
Public objPopturret As New clsPopTurret
Public objArm As New clsArm
Public objGasbag As New clsGasbag
'==================================================
' Class declarations for PSUEDO SUBS in toolbar
'==================================================
Public objCargo As New clsCargo
Public objHardpoint As New clsHardPoint
Public objArmor As New clsArmor
'==================================================
' Class declarations for Manned Vehicle Components
'==================================================
Public objManeuverControl As New clsManeuverControl
Public objCrewStation As New clsCrewStation
Public objAccommodation As New clsAccommodation
Public objEnvironmentalSystem As New clsEnvironmentalSystem
Public objSafetySystem As New clsSafetySystem
'==================================================
' Class declarations for Power and Fuel Components
'==================================================
Public objMuscleEngine As New clsMuscleEngine
Public objInternalCombustionEngine As New clsInternalCombustionEngine
Public objSteamEngine As New clsSteamEngine
Public objGasandMHDTurbine As New clsGasandMHDTurbine
Public objFuelCell As New clsFuelCell
Public objReactor As New clsReactor
Public objExoticPowerPlant As New clsExoticPowerPlant
Public objEnergyBank As New clsEnergyBank
Public objFuelBunker As New clsFuelBunker
Public objFuelTank As New clsFuelTank
Public objFuel As New clsFuel
Public objSnorkel As New clsSnorkel
Public objBeamedPowerReceiver As New clsBeamedPowerReceiver
Public objNitrousOxideBooster As New clsNitrousOxideBooster
'==================================================
' Class declarations for Instruments and Electronics
'==================================================
Public objCommunicator As New clsCommunicator
Public objSearchlight As New clsSearchlight
Public objVisualAugmentationSystem As New clsVisualAugmentationSystem
Public objExtendableSensorPeriscope As New clsExtendableSensorPeriscope
Public objRadarandLadar As New clsRadarandLadar
Public objSonar As New clsSonar
Public objThermPassElectromag As New clsThermPassElectromag
Public objOtherSensor As New clsOtherSensor
Public objSoundDetector As New clsSoundDetector
Public objScientificSensor As New clsScientificSensor
Public objAudioVisualSystem As New clsAudioVisualSystem
Public objNavigationSystem As New clsNavigationSystem
Public objTargetingSystem As New clsTargetingSystem
Public objElectronicCountermeasure As New clsElectronicCountermeasure
Public objComputer As New clsComputer
Public objTerminal As New clsTerminal
Public objSoftware As New clsSoftware
Public objNeuralInterfaceSystem As New clsNeuralInterfaceSystem
'==================================================
' Class declarations for Miscelleneous Equiptment
'==================================================
Public objArmMotor As New clsArmMotor
Public objFireExtinguisher As New clsFireExtinguisher
Public objBilgePump As New clsBilgePump
Public objLabandWorkshop As New clsLabandWorkshop
Public objHeavyEquiptment As New clsHeavyEquiptment
Public objEmergencyMedicalEquiptment As New clsEmergencyMedicalEquiptment
Public objEntertainmentFacility As New clsEntertainmentFacility
Public objVehicleAccess As New clsVehicleAccess
Public objTeleportProjector As New clsTeleportProjector
Public objSecurityDirtyTrick As New clsSecurityDirtyTrick
Public objVehicleStorage As New clsVehicleStorage
Public objLandingAid As New clsLandingAid
Public objFuelAccessory As New clsFuelAccessory
Public objNuclearDamper As New clsNuclearDamper
Public objRealityStabilizer As New clsRealityStabilizer
Public objModularSocket As New clsModularSocket
Public objModule As New clsModule

'==================================================
' Class declarations for WEAPONS and Weapon Accessories
'==================================================
Public objArtilleryandGun As New clsArtilleryandGun
Public objBeamWeapon As New clsBeamWeapon
Public objLauncher As New clsLauncher
Public objTorpedoandMissile As New clsTorpedoandMissile
Public objWeaponLink As New clsWeaponLink
Public objStabilizationGear As New clsStabilizationGear
Public objWeaponMount As New clsWeaponMount
Public objWeaponBay As New clsWeaponBay
Public objAntiBlastMagazine As New clsAntiBlastMagazine
'=====================================================
'Class declarations for drivetrains
'=====================================================
Public objWheeledDrivetrain As New clsWheeledDrivetrain
Public objTrackedDrivetrain As New clsTrackedDrivetrain
Public objLegDrivetrain As New clsLegDrivetrain
Public objFlexibodyDrivetrain As New clsFlexibodyDrivetrain
Public objOrnithopterDrivetrain As New clsOrnithopterDrivetrain
Public objTTRRotarDrivetrain As New clsTTRRotarDrivetrain
Public objMMRRotarDrivetrain As New clsMMRRotarDrivetrain
Public objCARRotarDrivetrain As New clsCARRotarDrivetrain
Public objAerialPropeller As New clsAerialPropeller
Public objDuctedFan As New clsDuctedFan
Public objAquaticPropulsion As New clsAquaticPropulsion
Public objHarness As New clsHarness
Public objMagLevLifter As New clsMagLevLifter
Public objJetEngine As New clsJetEngine
Public objReactionlessThruster As New clsReactionlessThruster
Public objRocketEngine As New clsRocketEngine
Public objRowingPositions As New clsRowingPositions
Public objSail As New clsSail
Public objAerialSail As New clsAerialSail
Public objLightSail As New clsLightSail
Public objStarDrive As New clsStardrive
Public objTeleportationDrive As New clsTeleportationDrive
Public objParachronicConveyer As New clsParachronicConveyer
Public objOrionEngine As New clsOrionEngine
Public objContraGravGenerator As New clsContraGravGenerator
Public objLiftingGas As New clsLiftingGas
'=====================================================
'=====================================================
'=====================================================
'=====================================================
' Constant declaration for all vehicle objects
'=====================================================
' Subassemblies
Public Const Body = 1
Public Const Skid = 2
Public Const Wheel = 3
Public Const Track = 4
Public Const Leg = 5
Public Const Wing = 6
Public Const Rotar = 7
Public Const Hydrofoil = 8
Public Const Hovercraft = 9
Public Const Superstructure = 10
Public Const OpenMount = 11
Public Const Mast = 12
Public Const Pod = 13
Public Const Turret = 14
Public Const Popturret = 15
Public Const Gasbag = 16
Public Const Arm = 17
' Other Toolbar Pseudo Subs
Public Const Armor = 18
Public Const Cargo = 23
Public Const Hardpoint = 26
'Instruments and Electronics
Public Const Communicator = 501
Public Const Searchlight = 502
Public Const VisualAugmentationSystem = 503
Public Const ExtendableSensorPeriscope = 504
Public Const RadarandLadar = 505
Public Const Sonar = 506
Public Const ThermPassElectromag = 507
Public Const OtherSensor = 508
Public Const SoundDetector = 509
Public Const ScientificSensor = 510
Public Const AudioVisualSystem = 511
Public Const NavigationSystem = 512
Public Const TargetingSystem = 513
Public Const ElectronicCountermeasure = 514
Public Const Computer = 515
Public Const Terminal = 516
Public Const Software = 517
Public Const NeuralInterfaceSystem = 518
'Miscellaneous Equiptment
Const ArmMotor = 601
Const FireExtinguisher = 602
Const BilgePump = 603
Const LabandWorkshop = 604
Const HeavyEquiptment = 605
Const EmergencyMedicalEquiptment = 606
Const EntertainmentFacility = 607
Const VehicleAccess = 608
Const TeleportProjector = 609
Const SecurityDirtyTrick = 610
Const VehicleStorage = 611
Const LandingAid = 612
Const FuelAccessory = 613
Const NuclearDamper = 614
Const RealityStabilizer = 615
Const ModularSocket = 616
Const Module = 617
'Manned Vehicle Components
Public Const ManeuverControl = 301
Public Const CrewStation = 302
Public Const Accommodation = 303
Public Const EnvironmentalSystem = 304
Public Const SafetySystem = 305 ' this includes Grav compensators
'Power and Fuel
Public Const MuscleEngine = 401
Public Const InternalCombustionEngine = 402
Public Const SteamEngine = 403
Public Const GasandMHDTurbine = 404
Public Const FuelCell = 405
Public Const Reactor = 406
Public Const ExoticPowerPlant = 407
Public Const EnergyBank = 408
Public Const FuelBunker = 409
Public Const FuelTank = 410
Public Const Fuel = 411
Public Const Snorkel = 412
Public Const BeamedPowerReceiver = 413
Public Const NitrousOxideBooster = 414

'Weapons and Weapon Accessories
Public Const ArtilleryandGun = 101
Public Const BeamWeapon = 102
Public Const Launcher = 103
Public Const TorpedoandMissile = 104
Public Const WeaponLink = 105
Public Const StabilizationGear = 106
Public Const WeaponMount = 107
Public Const WeaponBay = 108
Public Const AntiBlastMagazine = 109
' Propulsion
Public Const WheeledDrivetrain = 201
Public Const TrackedDrivetrain = 202
Public Const LegDrivetrain = 203
Public Const FlexibodyDrivetrain = 204
Public Const TTRRotarDrivetrain = 205
Public Const MMRRotarDrivetrain = 206
Public Const CARRotarDrivetrain = 207
Public Const OrnithopterDrivetrain = 208
Public Const AirscrewDrivetrain = 209
Public Const AerialPropeller = 209
Public Const DuctedFan = 224
Public Const AquaticPropulsion = 210
Public Const Harness = 211
Public Const MagLevLifter = 212
Public Const JetEngine = 213
Public Const ReactionlessThruster = 214
Public Const RocketEngine = 215
Public Const RowingPositions = 216
Public Const Sail = 217
Public Const AerialSail = 218
Public Const LightSail = 219
Public Const StarDrive = 220
Public Const TeleportationDrive = 221
Public Const ParachronicConveyer = 222
Public Const OrionEngine = 223
Public Const ContraGravGenerator = 225
Public Const LiftingGas = 226

Sub Main()
Dim i As Long ' debug delay counter to keep the splash screen visible longer
    frmSplash.Show
    frmSplash.Refresh
    'Set fMainForm = New frmDesigner
    i = Timer + 1 ' 1 seconds
    Do While Timer < i
    Loop
    'Load fMainForm
    Load frmDesigner
    'load the armor table
    LoadArmorTable
    'load the surface table
    LoadSurfaceFeatureTable
    'load the ground stability/maneuver ratings table
    LoadGroundStabilityTable
    'load the water stability/maneuver ratings table
    LoadWaterStabilityTable
    Unload frmSplash
    'fMainForm.Show
    frmDesigner.Show
End Sub
Sub FillProfilesCombo()
Dim TempText As String
Dim TempText2 As String
Dim element As Object
'Save original text and clear the combos
TempText = frmProfiles.Combo1.Text
TempText2 = frmPerformance.Combo1.Text
frmProfiles.Combo1.Clear
frmPerformance.Combo1.Clear
frmProfiles.Combo1.Text = TempText
frmPerformance.Combo1.Text = TempText2
' fill the combo box
For Each element In colVehicle
    If TypeName(element) = "clsPerformance" Then
        frmProfiles.Combo1.AddItem element.Key
        frmPerformance.Combo1.AddItem element.Key
        If frmProfiles.Combo1.Text = "" Then
            frmProfiles.Combo1.Text = element.Key
        End If
        If frmPerformance.Combo1.Text = "" Then
            frmPerformance.Combo1.Text = element.Key
        End If
    End If
Next
End Sub

Sub FillProfilesList()
Dim element As Object
Dim itmX As ListItem
Dim i As Integer

If frmProfiles.Combo1.Text = "" Then Exit Sub
'clear the listview
frmProfiles.ListView1.ListItems.Clear
' cycle through all the objects
For Each element In colVehicle
    ' only pick those elements that fit into that performance category
    Select Case TypeName(colVehicle.Item(colVehicle.Item(frmProfiles.Combo1.Text).PerformanceType))
        Case "clsWheel"
            Select Case TypeName(element)
                Case "clsWheeledDrivetrain", "clsAnimal", "clsSail", "clsAirscrewDrivetrain", "clsJetEngine", "clsRocketEngine", "clsReactionlessThruster"
                    ' add the item to the listview
                    Set itmX = frmProfiles.ListView1.ListItems.Add()
                    'set icon to unchecked
                    itmX.SmallIcon = 1
                    itmX.Text = frmDesigner.TreeView1.Nodes(element.Key).Text
                    itmX.Key = element.Key
                    ' Add the item to the array within the performance class
                    colVehicle.Item(frmProfiles.Combo1.Text).Add element.Key, TypeName(element)
                    ' determine if the item should be checked or not
                    If colVehicle.Item(frmProfiles.Combo1.Text).Checked(element.Key) Then
                        itmX.SmallIcon = 2 ' check the item
                    End If
            End Select
    End Select
Next
End Sub

Sub FillNewProfilesList()

Dim element As Object
Dim itmX As ListItem
Dim i As Integer
Dim WaterPerformance As Boolean
Dim SubmergedPerformance As Boolean

'clear the listview
frmProfiles.ListView1.ListItems.Clear
' cycle through all the objects
For Each element In colVehicle
    ' only pick Motive Assemblies and Flexibody and Water or Submerged type stuff
            Select Case TypeName(element)
                Case "clsWheel", "clsTrack", "clsSkid", "clsFlexibodyDrivetrain"
                    ' add the item to the listview
                    Set itmX = frmNewProfile.ListView1.ListItems.Add()
                    'set icon to unchecked
                    itmX.SmallIcon = 1
                    itmX.Text = frmDesigner.TreeView1.Nodes(element.Key).Text
                    itmX.Key = element.Key
            End Select
Next
' add items for Water performance if applicable
If colVehicle("1_").WaterProof Then WaterPerformance = True
If colVehicle("1_").Submersible Then
    WaterPerformance = True
    SubmergedPerformance = True
End If
If colVehicle("1_").Sealed Then WaterPerformance = True
If WaterPerformance Then
    ' add the item to the listview
    Set itmX = frmNewProfile.ListView1.ListItems.Add()
    'set icon to unchecked
    itmX.SmallIcon = 1
    itmX.Text = "Water Performance"
    itmX.Key = "Water__"
End If
If SubmergedPerformance = True Then
    ' add the item to the listview
    Set itmX = frmNewProfile.ListView1.ListItems.Add()
    'set icon to unchecked
    itmX.SmallIcon = 1
    itmX.Text = "Submerged Performance"
    itmX.Key = "Submerged__"
End If

End Sub

Function CheckValidLocation(Key As String, ParentKey As String) As Boolean
' Checks to see if the user is allowed to drop the node in the specified location
    CheckValidLocation = True

End Function

Function CheckMinTL(TL As Integer) As Boolean

If gVehicleTL < TL Then
    MsgBox "This option cannot be set.  Vehicle Tech Level must be " & TL & " or higher to enable this option.", vbOKOnly, "Tech Level too low!"
    CheckMinTL = False
Else
    CheckMinTL = True
End If

End Function


Sub LoadListView(TableName As String)
Dim db As Database
Dim ws As Workspace
Dim rs As Recordset
Dim ctI As Integer
Dim ctJ As Integer
    ' Clear the listview
    frmDesigner.ListView1.ListItems.Clear
    frmDesigner.ListView1.ColumnHeaders.Clear
    ' Set listbox to report mode
    frmDesigner.ListView1.View = lvwReport

    ' Open a workspace, database and recordset
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\gurps.mdb")
    Set rs = db.OpenRecordset(TableName, dbOpenTable)
    
    ' Load the Column Headers
    ctJ = 3 ' intialize counter
    For ctI = 1 To rs.Fields.Count - 3
        frmDesigner.ListView1.ColumnHeaders.Add ctI, , rs.Fields(ctJ).Name
        ctJ = ctJ + 1
    Next
    rs.MoveFirst 'move to first recordset
    ' Load the fields from the database
    ctJ = 1 ' initialize counter
    While rs.EOF <> True
        frmDesigner.ListView1.ListItems.Add , , rs.Fields(3).Value, , rs.Fields(1).Value
        For ctI = 1 To rs.Fields.Count - 4
             frmDesigner.ListView1.ListItems.Item(ctJ).SubItems(ctI) = rs.Fields(ctI + 3).Value
        Next
        ctJ = ctJ + 1 'incriment counter
        rs.MoveNext ' move to next recordset
    Wend
    rs.Close
    db.Close
    ws.Close
End Sub

Sub GetDataTypeandIcons(SelectedItem As String, TableName As String)
Dim db As Database
Dim ws As Workspace
Dim rs As Recordset
Dim TempType As Integer
Dim Icon1 As Integer
Dim Icon2 As Integer

' Open a workspace, database and recordset
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(App.Path & "\gurps.mdb")
    Set rs = db.OpenRecordset(TableName, dbOpenTable)
' Move to the first recordset
    rs.MoveFirst
' Loop thru all records until you find the "Selected Item"
    While rs.EOF <> True
        If rs.Fields(3) = SelectedItem Then
            TempType = rs.Fields(0).Value
            Icon1 = rs.Fields(1).Value
            Icon2 = rs.Fields(2).Value
        End If
        rs.MoveNext ' move to next recordset
    Wend
' Close the database, workspace and recordset
    rs.Close
    db.Close
    ws.Close
    
' Update the variables
myImage = Icon1
mySelectedImage = Icon2
Datatype = TempType

End Sub

Sub AddNewChildNode(NodeName, Image, SelectedImage, sKey)
'Add a node using tvwChild
    Dim oNodex As node
    'Dim skey As String
    Dim iIndex As Integer
        
    Set frmDesigner.TreeView1.ImageList = frmDesigner.ImageList1
        
    On Error GoTo myerr 'if the treeview does not have a node selected
    ' the next line of code will return an error number 91
    iIndex = frmDesigner.TreeView1.DropHighlight.Index 'Check to see if a Node is selected
    'skey = GetNextKey() ' Get a key for the new Node
    Set oNodex = frmDesigner.TreeView1.Nodes.Add(iIndex, tvwChild, sKey, NodeName, Image, SelectedImage)
    oNodex.EnsureVisible 'make sure the child node is visible
    Exit Sub
myerr:
    'Display a messge telling the user to select a node
    MsgBox ("You must select a Node to do an Add Child" & vbCrLf _
       & "If the TreeView is empty us Add Last to create the first node")
    Exit Sub
End Sub

Sub GetFirstParent()
'Find the first parent node in the TreeView
On Error GoTo myerr
Dim i As Integer
Dim nTmp As Integer
    For i = 1 To frmDesigner.TreeView1.Nodes.Count
        'This will give an error if there is no parent
        nTmp = frmDesigner.TreeView1.Nodes(i).Parent.Index
     Next
    Exit Sub
myerr:
    mnIndex = i
    Exit Sub
End Sub

Function GetNextKey() As String
'Returns a new key value for each Node being added to the TreeView
'This algorithm is very simple and will limit you to adding a total of 999 nodes
'Each node needs a unique key and to allow removing nodes you can't use the Nodes count +1
'as the key for a new node.
    Dim sNewKey As String
    Dim iHold As Integer
    Dim i As Integer
    On Error GoTo myerr
    'The next line will return error #35600 if there are no Nodes in the TreeView
    iHold = Val(frmDesigner.TreeView1.Nodes(1).Key)
    For i = 1 To frmDesigner.TreeView1.Nodes.Count
        If Val(frmDesigner.TreeView1.Nodes(i).Key) > iHold Then
            iHold = Val(frmDesigner.TreeView1.Nodes(i).Key)
        End If
    Next
    iHold = iHold + 1
    sNewKey = CStr(iHold) & "_"
    GetNextKey = sNewKey 'Return a unique key
    Exit Function
myerr:
    'Because the TreeView is empty return a 1 for the key of the first Node
    'DEBUG MsgBox Error & "  This is a handled error in Function GetNextKey."
    GetNextKey = "1_"
    Exit Function
End Function


Sub RemoveNode(Key As String)
    'Remove the selected Node and the Collection item
    'Dim sKey As String
    
    'On Error GoTo myerr 'if the treeview does not have a node selected
    ' the next line of code will return an error number 91
    'iIndex = frmDesigner.TreeView1.SelectedItem.Index 'Check to see if a Node is selected
    
    
    'If the Node has Children call the sub that writes the children
    If frmDesigner.TreeView1.Nodes(Key).Children > 0 Then
        RemoveChild frmDesigner.TreeView1.Nodes(Key).Index
    End If
    
    frmDesigner.TreeView1.Nodes.Remove Key 'Removes the Node and any children it has
    colVehicle.Remove Key ' remove the item from the collection
    
    'Exit Sub
'myerr:
    'Display a messge telling the user to select a node
 '   MsgBox ("Nothing to delete")
    'Exit Sub
End Sub
    
Sub RemoveChild(iNodeIndex As Integer)
' Write the child nodes to the table. This sub uses recursion
' to loop through the child nodes. It receives the Index of
' the node that has the children
    Dim i As Integer
    Dim iTempIndex As Integer
    iTempIndex = frmDesigner.TreeView1.Nodes(iNodeIndex).Child.FirstSibling.Index
    'Loop through all a Parents Child Nodes
    For i = 1 To frmDesigner.TreeView1.Nodes(iNodeIndex).Children
        colVehicle.Remove frmDesigner.TreeView1.Nodes(iTempIndex).Key
    
    ' If the Node we are on has a child call the Sub again
    If frmDesigner.TreeView1.Nodes(iTempIndex).Children > 0 Then
       RemoveChild (iTempIndex)
    End If
    
  ' If we are not on the last child move to the next child Node
    If i <> frmDesigner.TreeView1.Nodes(iNodeIndex).Children Then
       iTempIndex = frmDesigner.TreeView1.Nodes(iTempIndex).Next.Index
    End If
  Next i
End Sub
    
Sub UnloadAllForms(sFormName As String)

Dim Form As Form
   For Each Form In Forms
      If Form.Name <> sFormName Then
         Unload Form
         Set Form = Nothing
      End If
   Next Form
End Sub

Sub AddObject(Datatype As Integer, CurrentItem As String, arrProperties() As String)
            
On Error GoTo ErrorHandler2

Select Case Datatype
'//////////////////////////////////////////////////////////////////
'Weapon and Weapon Accessories
'//////////////////////////////////////////////////////////////////
Case ArtilleryandGun
   
    Set objArtilleryandGun = New clsArtilleryandGun
    colVehicle.Add objArtilleryandGun, CurrentItem
     
     With colVehicle.Item(CurrentItem) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage

        '.Name = arrProperties(1)
        .Malf = arrProperties(3)
        .objType = arrProperties(4)
        .Damage = arrProperties(5)
        .SS = arrProperties(6)
        .Acc = arrProperties(7)
        .halfD = arrProperties(8)
        .Max = arrProperties(9)
        .RoF = arrProperties(10)
        .Weight = Val(arrProperties(11))
        .Cost = arrProperties(12)
        .WPS = arrProperties(13)
        .VPS = arrProperties(14)
        .CPS = arrProperties(15)
        .Ldrs = arrProperties(16)
   End With
 
Case BeamWeapon
    Set objBeamWeapon = New clsBeamWeapon
    colVehicle.Add objBeamWeapon, CurrentItem
    With colVehicle.Item(CurrentItem) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With

Case Launcher
    Set objLauncher = New clsLauncher
    colVehicle.Add objLauncher, CurrentItem
    With colVehicle.Item(CurrentItem) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case TorpedoandMissile
    Set objTorpedoandMissile = New clsTorpedoandMissile
    colVehicle.Add objTorpedoandMissile, CurrentItem
    With colVehicle.Item(CurrentItem) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case WeaponLink
    Set objWeaponLink = New clsWeaponLink
    colVehicle.Add objWeaponLink, CurrentItem
    With colVehicle.Item(CurrentItem) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case StabilizationGear
    Set objStabilizationGear = New clsStabilizationGear
    colVehicle.Add objStabilizationGear, CurrentItem
    With colVehicle.Item(CurrentItem) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case WeaponMount
    Set objWeaponMount = New clsWeaponMount
    colVehicle.Add objWeaponMount, CurrentItem
    With colVehicle.Item(CurrentItem) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case WeaponBay
    Set objWeaponBay = New clsWeaponBay
    colVehicle.Add objWeaponBay, CurrentItem
    With colVehicle.Item(CurrentItem) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case AntiBlastMagazine
    Set objAntiBlastMagazine = New clsAntiBlastMagazine
    colVehicle.Add objAntiBlastMagazine, CurrentItem
    With colVehicle.Item(CurrentItem) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
'/////////////////////////////////////////////////////////////////
'Propulsion Systems
'/////////////////////////////////////////////////////////////////
Case WheeledDrivetrain
    Set objWheeledDrivetrain = New clsWheeledDrivetrain
    colVehicle.Add objWheeledDrivetrain, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case TrackedDrivetrain
    Set objTrackedDrivetrain = New clsTrackedDrivetrain
    colVehicle.Add objTrackedDrivetrain, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case LegDrivetrain
    Set objLegDrivetrain = New clsLegDrivetrain
    colVehicle.Add objLegDrivetrain, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case FlexibodyDrivetrain
    Set objFlexibodyDrivetrain = New clsFlexibodyDrivetrain
    colVehicle.Add objFlexibodyDrivetrain, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case TTRRotarDrivetrain
    Set objTTRRotarDrivetrain = New clsTTRRotarDrivetrain
    colVehicle.Add objTTRRotarDrivetrain, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case MMRRotarDrivetrain
    Set objMMRRotarDrivetrain = New clsMMRRotarDrivetrain
    colVehicle.Add objMMRRotarDrivetrain, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case CARRotarDrivetrain
    Set objCARRotarDrivetrain = New clsCARRotarDrivetrain
    colVehicle.Add objCARRotarDrivetrain, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case OrnithopterDrivetrain
    Set objOrnithopterDrivetrain = New clsOrnithopterDrivetrain
    colVehicle.Add objOrnithopterDrivetrain, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case AerialPropeller
    Set objAerialPropeller = New clsAerialPropeller
    colVehicle.Add objAerialPropeller, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case DuctedFan
    Set objDuctedFan = New clsDuctedFan
    colVehicle.Add objDuctedFan, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case AquaticPropulsion
    Set objAquaticPropulsion = New clsAquaticPropulsion
    colVehicle.Add objAquaticPropulsion, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case Harness
    Set objHarness = New clsHarness
    colVehicle.Add objHarness, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case MagLevLifter
    Set objMagLevLifter = New clsMagLevLifter
    colVehicle.Add objMagLevLifter, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case JetEngine
    Set objJetEngine = New clsJetEngine
    colVehicle.Add objJetEngine, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case ReactionlessThruster
    Set objReactionlessThruster = New clsReactionlessThruster
    colVehicle.Add objReactionlessThruster, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case RocketEngine
    Set objRocketEngine = New clsRocketEngine
    colVehicle.Add objRocketEngine, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case RowingPositions
    Set objRowingPositions = New clsRowingPositions
    colVehicle.Add objRowingPositions, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case Sail
    Set objSail = New clsSail
    colVehicle.Add objSail, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case AerialSail
    Set objAerialSail = New clsAerialSail
    colVehicle.Add objAerialSail, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case LightSail
    Set objLightSail = New clsLightSail
    colVehicle.Add objLightSail, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case StarDrive
    Set objStarDrive = New clsStardrive
    colVehicle.Add objStarDrive, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case OrionEngine
    Set objOrionEngine = New clsOrionEngine
    colVehicle.Add objOrionEngine, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case TeleportationDrive
    Set objTeleportationDrive = New clsTeleportationDrive
    colVehicle.Add objTeleportationDrive, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case ParachronicConveyer
    Set objParachronicConveyer = New clsParachronicConveyer
    colVehicle.Add objParachronicConveyer, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
'/////////////////////////////////////////////////////////////////
'Aerostatic Lift Systems
'/////////////////////////////////////////////////////////////////
Case ContraGravGenerator
    Set objContraGravGenerator = New clsContraGravGenerator
    colVehicle.Add objContraGravGenerator, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case LiftingGas
    Set objLiftingGas = New clsLiftingGas
    colVehicle.Add objLiftingGas, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
'/////////////////////////////////////////////////////////////////
'Instruments and Electronics
'/////////////////////////////////////////////////////////////////
Case Communicator
    Set objCommunicator = New clsCommunicator
    colVehicle.Add objCommunicator, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case Searchlight
    Set objSearchlight = New clsSearchlight
    colVehicle.Add objSearchlight, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case VisualAugmentationSystem
    Set objVisualAugmentationSystem = New clsVisualAugmentationSystem
    colVehicle.Add objVisualAugmentationSystem, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case ExtendableSensorPeriscope
    Set objExtendableSensorPeriscope = New clsExtendableSensorPeriscope
    colVehicle.Add objExtendableSensorPeriscope, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case RadarandLadar
    Set objRadarandLadar = New clsRadarandLadar
    colVehicle.Add objRadarandLadar, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case Sonar
    Set objSonar = New clsSonar
    colVehicle.Add objSonar, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case ThermPassElectromag
    Set objThermPassElectromag = New clsThermPassElectromag
    colVehicle.Add objThermPassElectromag, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case OtherSensor
    Set objOtherSensor = New clsOtherSensor
    colVehicle.Add objOtherSensor, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case SoundDetector
    Set objSoundDetector = New clsSoundDetector
    colVehicle.Add objSoundDetector, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case ScientificSensor
    Set objScientificSensor = New clsScientificSensor
    colVehicle.Add objScientificSensor, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case AudioVisualSystem
    Set objAudioVisualSystem = New clsAudioVisualSystem
    colVehicle.Add objAudioVisualSystem, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case NavigationSystem
    Set objNavigationSystem = New clsNavigationSystem
    colVehicle.Add objNavigationSystem, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case TargetingSystem
    Set objTargetingSystem = New clsTargetingSystem
    colVehicle.Add objTargetingSystem, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case ElectronicCountermeasure
    Set objElectronicCountermeasure = New clsElectronicCountermeasure
    colVehicle.Add objElectronicCountermeasure, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case Computer
    Set objComputer = New clsComputer
    colVehicle.Add objComputer, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case Terminal
    Set objTerminal = New clsTerminal
    colVehicle.Add objTerminal, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case Software
    Set objSoftware = New clsSoftware
    colVehicle.Add objSoftware, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case NeuralInterfaceSystem
    Set objNeuralInterfaceSystem = New clsNeuralInterfaceSystem
    colVehicle.Add objNeuralInterfaceSystem, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
'/////////////////////////////////////////////////////////////////
'Miscellaneous Equiptment
'/////////////////////////////////////////////////////////////////
Case ArmMotor
Set objArmMotor = New clsArmMotor
    colVehicle.Add objArmMotor, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case FireExtinguisher
    Set objFireExtinguisher = New clsFireExtinguisher
    colVehicle.Add objFireExtinguisher, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case BilgePump
    Set objBilgePump = New clsBilgePump
    colVehicle.Add objBilgePump, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case LabandWorkshop
    Set objLabandWorkshop = New clsLabandWorkshop
    colVehicle.Add objLabandWorkshop, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case HeavyEquiptment
    Set objHeavyEquiptment = New clsHeavyEquiptment
    colVehicle.Add objHeavyEquiptment, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case EmergencyMedicalEquiptment
    Set objEmergencyMedicalEquiptment = New clsEmergencyMedicalEquiptment
    colVehicle.Add objEmergencyMedicalEquiptment, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case EntertainmentFacility
    Set objEntertainmentFacility = New clsEntertainmentFacility
    colVehicle.Add objEntertainmentFacility, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case VehicleAccess
    Set objVehicleAccess = New clsVehicleAccess
    colVehicle.Add objVehicleAccess, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case TeleportProjector
    Set objTeleportProjector = New clsTeleportProjector
    colVehicle.Add objTeleportProjector, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case SecurityDirtyTrick
    Set objSecurityDirtyTrick = New clsSecurityDirtyTrick
    colVehicle.Add objSecurityDirtyTrick, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case VehicleStorage
    Set objVehicleStorage = New clsVehicleStorage
    colVehicle.Add objVehicleStorage, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case LandingAid
    Set objLandingAid = New clsLandingAid
    colVehicle.Add objLandingAid, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case FuelAccessory
    Set objFuelAccessory = New clsFuelAccessory
    colVehicle.Add objFuelAccessory, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case NuclearDamper
    Set objNuclearDamper = New clsNuclearDamper
    colVehicle.Add objNuclearDamper, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case RealityStabilizer
    Set objRealityStabilizer = New clsRealityStabilizer
    colVehicle.Add objRealityStabilizer, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case ModularSocket
    Set objModularSocket = New clsModularSocket
    colVehicle.Add objModularSocket, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case Module
    Set objModule = New clsModule
    colVehicle.Add objModule, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
'/////////////////////////////////////////////////////////////////
'Crew, Passengers, Seating, Seats, Stations
'/////////////////////////////////////////////////////////////////
Case ManeuverControl
    Set objManeuverControl = New clsManeuverControl
    colVehicle.Add objManeuverControl, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case CrewStation
    Set objCrewStation = New clsCrewStation
    colVehicle.Add objCrewStation, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case Accommodation
    Set objAccommodation = New clsAccommodation
    colVehicle.Add objAccommodation, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case EnvironmentalSystem
    Set objEnvironmentalSystem = New clsEnvironmentalSystem
    colVehicle.Add objEnvironmentalSystem, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case SafetySystem
    Set objSafetySystem = New clsSafetySystem
    colVehicle.Add objSafetySystem, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
'/////////////////////////////////////////////////////////////////
'Power and Fuel Components
'/////////////////////////////////////////////////////////////////
Case MuscleEngine
    Set objMuscleEngine = New clsMuscleEngine
    colVehicle.Add objMuscleEngine, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case InternalCombustionEngine
    Set objInternalCombustionEngine = New clsInternalCombustionEngine
    colVehicle.Add objInternalCombustionEngine, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case SteamEngine
    Set objSteamEngine = New clsSteamEngine
    colVehicle.Add objSteamEngine, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case GasandMHDTurbine
    Set objGasandMHDTurbine = New clsGasandMHDTurbine
    colVehicle.Add objGasandMHDTurbine, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case FuelCell
    Set objFuelCell = New clsFuelCell
    colVehicle.Add objFuelCell, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case Reactor
    Set objReactor = New clsReactor
    colVehicle.Add objReactor, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case ExoticPowerPlant
    Set objExoticPowerPlant = New clsExoticPowerPlant
    colVehicle.Add objExoticPowerPlant, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case EnergyBank
    Set objEnergyBank = New clsEnergyBank
    colVehicle.Add objEnergyBank, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case FuelBunker
    Set objFuelBunker = New clsFuelBunker
    colVehicle.Add objFuelBunker, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case FuelTank
    Set objFuelTank = New clsFuelTank
    colVehicle.Add objFuelTank, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case Fuel
    Set objFuel = New clsFuel
    colVehicle.Add objFuel, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case Snorkel
    Set objSnorkel = New clsSnorkel
    colVehicle.Add objSnorkel, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case BeamedPowerReceiver
    Set objBeamedPowerReceiver = New clsBeamedPowerReceiver
    colVehicle.Add objBeamedPowerReceiver, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
Case NitrousOxideBooster
    Set objNitrousOxideBooster = New clsNitrousOxideBooster
    colVehicle.Add objNitrousOxideBooster, CurrentItem
    With colVehicle.Item(CurrentItem)
        .Key = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Key
        .Parent = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Parent.Key ' root node doesnt have a parent right?
        .Text = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Text
        .Image = frmDesigner.TreeView1.Nodes.Item(CurrentItem).Image
        .SelectedImage = frmDesigner.TreeView1.Nodes.Item(CurrentItem).SelectedImage
    End With
End Select
ErrorHandler2:
    ' if the nodes parent property isnt set, that means its a root node so skip
    ' the line that tries to set the .parent from the nodes non existant parent property
   'MsgBox Error
 Resume Next
End Sub

Public Sub Properties_Show(Key As String)
' to increase performance, could possible put giant WITH block in

Dim ComponentsParent As String
Dim component As String

If Key <> "1_" Then
    ComponentsParent = TypeName(colVehicle.Item((colVehicle.Item(Key).Parent)))
End If

component = TypeName(colVehicle.Item(Key))

' set the # of rows in the grid
frmProperties.MSFlexGrid1.Rows = 0

With colVehicle.Item(Key)
' Fill the window with properties for the correct Collection item
Select Case component
        
        Case "clsArtilleryandGun"
            
'            frmProperties.MSFlexGrid1.AddItem "Tech Level" & vbTab & .TL
            frmProperties.MSFlexGrid1.AddItem "Text" & vbTab & .Text
            frmProperties.MSFlexGrid1.AddItem "Sub-Type" & vbTab & .SubType
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & Str(.Cost)
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & Str(.Weight)
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & Str(.Volume)
'            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & Str(.SurfaceArea)
'           frmProperties.MSFlexGrid1.AddItem "Frame Strength" & vbTab & .FrameStrength
'            frmProperties.MSFlexGrid1.AddItem "Materials" & vbTab & .Materials
'            frmProperties.MSFlexGrid1.AddItem "Responsive" & vbTab & .Responsive
'            frmProperties.MSFlexGrid1.AddItem "Robotic" & vbTab & .Robotic
'            frmProperties.MSFlexGrid1.AddItem "Biomechanical" & vbTab & .Biomechanical
'            frmProperties.MSFlexGrid1.AddItem "Living Metal" & vbTab & .LivingMetal
            frmProperties.Text1.Tag = Key
            
        Case "clsWheel"
            frmProperties.MSFlexGrid1.AddItem "Tech Level" & vbTab & .TL
            frmProperties.MSFlexGrid1.AddItem "Sub-Type" & vbTab & .SubType
            frmProperties.MSFlexGrid1.AddItem "Number of Wheels" & vbTab & .Number
            frmProperties.MSFlexGrid1.AddItem "Improved Suspension" & vbTab & .ImprovedSuspension
            frmProperties.MSFlexGrid1.AddItem "Wheel Blades" & vbTab & .Wheelblades
            frmProperties.MSFlexGrid1.AddItem "Snow Tires" & vbTab & .SnowTires
            frmProperties.MSFlexGrid1.AddItem "Racing Tires" & vbTab & .RacingTires
            frmProperties.MSFlexGrid1.AddItem "Puncture Resistant" & vbTab & .PunctureResistant
            frmProperties.MSFlexGrid1.AddItem "Retract Location" & vbTab & .RetractLocation
            frmProperties.MSFlexGrid1.AddItem "Improved Breaks" & vbTab & .ImprovedBreaks
            frmProperties.MSFlexGrid1.AddItem "All Wheel Steering" & vbTab & .AllwheelSteering
            frmProperties.MSFlexGrid1.AddItem "Smart Wheels" & vbTab & .Smartwheels
            frmProperties.MSFlexGrid1.AddItem "Frame Strength" & vbTab & .FrameStrength
            frmProperties.MSFlexGrid1.AddItem "Materials" & vbTab & .Materials
            frmProperties.MSFlexGrid1.AddItem "Responsive" & vbTab & .Responsive
            frmProperties.MSFlexGrid1.AddItem "Robotic" & vbTab & .Robotic
            frmProperties.MSFlexGrid1.AddItem "Biomechanical" & vbTab & .Biomechanical
            frmProperties.MSFlexGrid1.AddItem "Living Metal" & vbTab & .LivingMetal
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & Str(.Cost)
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & Str(.Weight)
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & Str(.SurfaceArea)
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & Str(.Volume)
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & Str(.HitPoints)
            frmProperties.Text1.Tag = Key
         
        Case "clsBody"
            frmProperties.MSFlexGrid1.AddItem "Slope Right" & vbTab & .SlopeR
            frmProperties.MSFlexGrid1.AddItem "Slope Left" & vbTab & .SlopeL
            frmProperties.MSFlexGrid1.AddItem "Slope Front" & vbTab & .SlopeF
            frmProperties.MSFlexGrid1.AddItem "Slope Back" & vbTab & .SlopeB
            frmProperties.MSFlexGrid1.AddItem "NumLegs" & vbTab & .NumLegs
            frmProperties.MSFlexGrid1.AddItem "Wings or Rotars" & vbTab & .WingsorRotors
            frmProperties.MSFlexGrid1.AddItem "Improved Suspension" & vbTab & .ImprovedSuspension
            frmProperties.MSFlexGrid1.AddItem "Retractable Wheels / Skids" & vbTab & .RetractableWheelsSkids
            frmProperties.MSFlexGrid1.AddItem "Tech Level" & vbTab & .TL
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & Str(.Cost)
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & Str(.Weight)
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & Str(.Volume)
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & Str(.SurfaceArea)
            frmProperties.MSFlexGrid1.AddItem "TotalRotation Space" & vbTab & .TotalRotationSpace
            frmProperties.MSFlexGrid1.AddItem "Minimum Volume" & vbTab & .MinimumVolume
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.MSFlexGrid1.AddItem "Flotation Rating" & vbTab & .FloatationRating
            frmProperties.MSFlexGrid1.AddItem "Total Surface Area" & vbTab & .TotalSurfaceArea
            frmProperties.MSFlexGrid1.AddItem "Structural Surface Area" & vbTab & .StructuralSurfaceArea
            frmProperties.MSFlexGrid1.AddItem "Frame Strength" & vbTab & .FrameStrength
            frmProperties.MSFlexGrid1.AddItem "Materials" & vbTab & .Materials
            frmProperties.MSFlexGrid1.AddItem "Responsive" & vbTab & .Responsive
            frmProperties.MSFlexGrid1.AddItem "Robotic" & vbTab & .Robotic
            frmProperties.MSFlexGrid1.AddItem "Biomechanical" & vbTab & .Biomechanical
            frmProperties.MSFlexGrid1.AddItem "Living Metal" & vbTab & .LivingMetal
            frmProperties.MSFlexGrid1.AddItem "Flotation Hull" & vbTab & .FloatationHull
            frmProperties.MSFlexGrid1.AddItem "Compartmentalization" & vbTab & .Compartmentalization
            frmProperties.MSFlexGrid1.AddItem "Combined Component Volume" & vbTab & .CombinedComponentVolume
            frmProperties.Text1.Tag = Key
        
        Case "clsSkid"
            frmProperties.MSFlexGrid1.AddItem "Tech Level" & vbTab & .TL
            frmProperties.MSFlexGrid1.AddItem "Number of Skids" & vbTab & .Number
            frmProperties.MSFlexGrid1.AddItem "Improved Suspension" & vbTab & .ImprovedSuspension
            frmProperties.MSFlexGrid1.AddItem "Retract Location" & vbTab & .RetractLocation
            frmProperties.MSFlexGrid1.AddItem "Frame Strength" & vbTab & .FrameStrength
            frmProperties.MSFlexGrid1.AddItem "Materials" & vbTab & .Materials
            frmProperties.MSFlexGrid1.AddItem "Responsive" & vbTab & .Responsive
            frmProperties.MSFlexGrid1.AddItem "Robotic" & vbTab & .Robotic
            frmProperties.MSFlexGrid1.AddItem "Biomechanical" & vbTab & .Biomechanical
            frmProperties.MSFlexGrid1.AddItem "Living Metal" & vbTab & .LivingMetal
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & Str(.Cost)
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & Str(.Weight)
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & Str(.Volume)
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & Str(.SurfaceArea)
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & Str(.HitPoints)
            frmProperties.Text1.Tag = Key
                
        Case "clsTrack"
            frmProperties.MSFlexGrid1.AddItem "Tech Level" & vbTab & .TL
            frmProperties.MSFlexGrid1.AddItem "Sub-Type" & vbTab & .SubType
            frmProperties.MSFlexGrid1.AddItem "Number of Tracks" & vbTab & .Number
            frmProperties.MSFlexGrid1.AddItem "Improved Suspension" & vbTab & .ImprovedSuspension
            frmProperties.MSFlexGrid1.AddItem "Frame Strength" & vbTab & .FrameStrength
            frmProperties.MSFlexGrid1.AddItem "Materials" & vbTab & .Materials
            frmProperties.MSFlexGrid1.AddItem "Responsive" & vbTab & .Responsive
            frmProperties.MSFlexGrid1.AddItem "Robotic" & vbTab & .Robotic
            frmProperties.MSFlexGrid1.AddItem "Biomechanical" & vbTab & .Biomechanical
            frmProperties.MSFlexGrid1.AddItem "Living Metal" & vbTab & .LivingMetal
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & Str(.Cost)
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & Str(.Weight)
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & Str(.Volume)
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & Str(.SurfaceArea)
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & Str(.HitPoints)
            frmProperties.Text1.Tag = Key
                                   
        
        Case "clsLeg"
            frmProperties.MSFlexGrid1.AddItem "Tech Level" & vbTab & .TL
            frmProperties.MSFlexGrid1.AddItem "Improved Suspension" & vbTab & .ImprovedSuspension
            frmProperties.MSFlexGrid1.AddItem "Frame Strength" & vbTab & .FrameStrength
            frmProperties.MSFlexGrid1.AddItem "Materials" & vbTab & .Materials
            frmProperties.MSFlexGrid1.AddItem "Responsive" & vbTab & .Responsive
            frmProperties.MSFlexGrid1.AddItem "Robotic" & vbTab & .Robotic
            frmProperties.MSFlexGrid1.AddItem "Biomechanical" & vbTab & .Biomechanical
            frmProperties.MSFlexGrid1.AddItem "Living Metal" & vbTab & .LivingMetal
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & Str(.Cost)
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & Str(.Weight)
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & Str(.Volume)
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & Str(.SurfaceArea)
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.Text1.Tag = Key
                                              
        Case "clsWing"
            frmProperties.MSFlexGrid1.AddItem "Tech Level" & vbTab & .TL
            frmProperties.MSFlexGrid1.AddItem "Sub-Type" & vbTab & .SubType
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & Str(.Cost)
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & Str(.Weight)
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & Str(.Volume)
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & Str(.SurfaceArea)
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.MSFlexGrid1.AddItem "Controlled Instability" & vbTab & .ControlledInstability
            frmProperties.MSFlexGrid1.AddItem "Folding Wings" & vbTab & .Folding
            frmProperties.MSFlexGrid1.AddItem "Variable Sweep Wings" & vbTab & .VariableSweep
            frmProperties.MSFlexGrid1.AddItem "Frame Strength" & vbTab & .FrameStrength
            frmProperties.MSFlexGrid1.AddItem "Materials" & vbTab & .Materials
            frmProperties.MSFlexGrid1.AddItem "Responsive" & vbTab & .Responsive
            frmProperties.MSFlexGrid1.AddItem "Robotic" & vbTab & .Robotic
            frmProperties.MSFlexGrid1.AddItem "Biomechanical" & vbTab & .Biomechanical
            frmProperties.MSFlexGrid1.AddItem "Living Metal" & vbTab & .LivingMetal
            frmProperties.Text1.Tag = Key
            
        Case "clsRotar"
            frmProperties.MSFlexGrid1.AddItem "Tech Level" & vbTab & .TL
            frmProperties.MSFlexGrid1.AddItem "Sub-Type" & vbTab & .SubType
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & Str(.Cost)
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & Str(.Weight)
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & Str(.Volume)
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & Str(.SurfaceArea)
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.MSFlexGrid1.AddItem "Controlled Instability" & vbTab & .ControlledInstability
            frmProperties.MSFlexGrid1.AddItem "Folding Rotars" & vbTab & .Folding
            frmProperties.MSFlexGrid1.AddItem "Number of Rotars" & vbTab & .Number
            frmProperties.MSFlexGrid1.AddItem "Frame Strength" & vbTab & .FrameStrength
            frmProperties.MSFlexGrid1.AddItem "Materials" & vbTab & .Materials
            frmProperties.MSFlexGrid1.AddItem "Responsive" & vbTab & .Responsive
            frmProperties.MSFlexGrid1.AddItem "Robotic" & vbTab & .Robotic
            frmProperties.MSFlexGrid1.AddItem "Biomechanical" & vbTab & .Biomechanical
            frmProperties.MSFlexGrid1.AddItem "Living Metal" & vbTab & .LivingMetal
            frmProperties.Text1.Tag = Key
            
        Case "clsHydrofoil"
                                   
            frmProperties.MSFlexGrid1.AddItem "Tech Level" & vbTab & .TL
            frmProperties.MSFlexGrid1.AddItem "Frame Strength" & vbTab & .FrameStrength
            frmProperties.MSFlexGrid1.AddItem "Materials" & vbTab & .Materials
            frmProperties.MSFlexGrid1.AddItem "Responsive" & vbTab & .Responsive
            frmProperties.MSFlexGrid1.AddItem "Robotic" & vbTab & .Robotic
            frmProperties.MSFlexGrid1.AddItem "Biomechanical" & vbTab & .Biomechanical
            frmProperties.MSFlexGrid1.AddItem "Living Metal" & vbTab & .LivingMetal
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & Str(.Cost)
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & Str(.Weight)
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & Str(.Volume)
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & Str(.SurfaceArea)
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.Text1.Tag = Key
            
        Case "clsHovercraft"
                                   
            frmProperties.MSFlexGrid1.AddItem "Tech Level" & vbTab & .TL
            frmProperties.MSFlexGrid1.AddItem "Sub-Type" & vbTab & .SubType
            frmProperties.MSFlexGrid1.AddItem "Frame Strength" & vbTab & .FrameStrength
            frmProperties.MSFlexGrid1.AddItem "Materials" & vbTab & .Materials
            frmProperties.MSFlexGrid1.AddItem "Responsive" & vbTab & .Responsive
            frmProperties.MSFlexGrid1.AddItem "Robotic" & vbTab & .Robotic
            frmProperties.MSFlexGrid1.AddItem "Biomechanical" & vbTab & .Biomechanical
            frmProperties.MSFlexGrid1.AddItem "Living Metal" & vbTab & .LivingMetal
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & Str(.Cost)
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & Str(.Weight)
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & Str(.Volume)
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & Str(.SurfaceArea)
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.Text1.Tag = Key
            
        Case "clsSuperStructure"
                                   
            frmProperties.MSFlexGrid1.AddItem "Tech Level" & vbTab & .TL
            frmProperties.MSFlexGrid1.AddItem "Slope Right" & vbTab & .SlopeR
            frmProperties.MSFlexGrid1.AddItem "Slope Left" & vbTab & .SlopeL
            frmProperties.MSFlexGrid1.AddItem "Slope Front" & vbTab & .SlopeF
            frmProperties.MSFlexGrid1.AddItem "Slope Back" & vbTab & .SlopeB
            frmProperties.MSFlexGrid1.AddItem "Compartmentalization" & vbTab & .Compartmentalization
            frmProperties.MSFlexGrid1.AddItem "Frame Strength" & vbTab & .FrameStrength
            frmProperties.MSFlexGrid1.AddItem "Materials" & vbTab & .Materials
            frmProperties.MSFlexGrid1.AddItem "Responsive" & vbTab & .Responsive
            frmProperties.MSFlexGrid1.AddItem "Robotic" & vbTab & .Robotic
            frmProperties.MSFlexGrid1.AddItem "Biomechanical" & vbTab & .Biomechanical
            frmProperties.MSFlexGrid1.AddItem "Living Metal" & vbTab & .LivingMetal
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & Str(.Cost)
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & Str(.Weight)
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & Str(.Volume)
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & Str(.SurfaceArea)
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.Text1.Tag = Key
            
        Case "clsOpenMount"
            frmProperties.MSFlexGrid1.AddItem "Tech Level" & vbTab & .TL
            frmProperties.MSFlexGrid1.AddItem "Text" & vbTab & .Text
            frmProperties.MSFlexGrid1.AddItem "Sub-Type" & vbTab & .SubType
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & Str(.Cost)
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & Str(.Weight)
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & Str(.Volume)
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & Str(.SurfaceArea)
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.MSFlexGrid1.AddItem "Rotation" & vbTab & .Rotation
            
            frmProperties.Text1.Tag = Key
            
        Case "clsMast"
            frmProperties.MSFlexGrid1.AddItem "Tech Level" & vbTab & .TL
            frmProperties.MSFlexGrid1.AddItem "Text" & vbTab & .Text
            frmProperties.MSFlexGrid1.AddItem "Sub-Type" & vbTab & .SubType
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & Str(.Cost)
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & Str(.Weight)
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & Str(.Volume)
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & Str(.SurfaceArea)
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & Str(.HitPoints)
            frmProperties.MSFlexGrid1.AddItem "Height" & vbTab & Str(.Height)
            frmProperties.MSFlexGrid1.AddItem "Materials" & vbTab & .Materials
            frmProperties.Text1.Tag = Key
                                                          
        Case "clsPod"
                                   
            frmProperties.MSFlexGrid1.AddItem "Tech Level" & vbTab & .TL
            frmProperties.MSFlexGrid1.AddItem "Frame Strength" & vbTab & .FrameStrength
            frmProperties.MSFlexGrid1.AddItem "Materials" & vbTab & .Materials
            frmProperties.MSFlexGrid1.AddItem "Responsive" & vbTab & .Responsive
            frmProperties.MSFlexGrid1.AddItem "Robotic" & vbTab & .Robotic
            frmProperties.MSFlexGrid1.AddItem "Biomechanical" & vbTab & .Biomechanical
            frmProperties.MSFlexGrid1.AddItem "Living Metal" & vbTab & .LivingMetal
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & Str(.Cost)
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & Str(.Weight)
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & Str(.Volume)
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & Str(.SurfaceArea)
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.Text1.Tag = Key
            
        Case "clsTurret"
                                   
            frmProperties.MSFlexGrid1.AddItem "Tech Level" & vbTab & .TL
            frmProperties.MSFlexGrid1.AddItem "Rotation" & vbTab & .Rotation
            frmProperties.MSFlexGrid1.AddItem "Slope Right" & vbTab & .SlopeR
            frmProperties.MSFlexGrid1.AddItem "Slope Left" & vbTab & .SlopeL
            frmProperties.MSFlexGrid1.AddItem "Slope Front" & vbTab & .SlopeF
            frmProperties.MSFlexGrid1.AddItem "Slope Back" & vbTab & .SlopeB
            frmProperties.MSFlexGrid1.AddItem "Compartmentalization" & vbTab & .Compartmentalization
            frmProperties.MSFlexGrid1.AddItem "Frame Strength" & vbTab & .FrameStrength
            frmProperties.MSFlexGrid1.AddItem "Materials" & vbTab & .Materials
            frmProperties.MSFlexGrid1.AddItem "Responsive" & vbTab & .Responsive
            frmProperties.MSFlexGrid1.AddItem "Robotic" & vbTab & .Robotic
            frmProperties.MSFlexGrid1.AddItem "Biomechanical" & vbTab & .Biomechanical
            frmProperties.MSFlexGrid1.AddItem "Living Metal" & vbTab & .LivingMetal
            frmProperties.MSFlexGrid1.AddItem "Rotation Space" & vbTab & .RotationSpace
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & Str(.Cost)
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & Str(.Weight)
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & Str(.Volume)
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & Str(.SurfaceArea)
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.Text1.Tag = Key
                
        Case "clsPopTurret"
                                   
            frmProperties.MSFlexGrid1.AddItem "Tech Level" & vbTab & .TL
            frmProperties.MSFlexGrid1.AddItem "Rotation" & vbTab & .Rotation
            frmProperties.MSFlexGrid1.AddItem "Slope Right" & vbTab & .SlopeR
            frmProperties.MSFlexGrid1.AddItem "Slope Left" & vbTab & .SlopeL
            frmProperties.MSFlexGrid1.AddItem "Slope Front" & vbTab & .SlopeF
            frmProperties.MSFlexGrid1.AddItem "Slope Back" & vbTab & .SlopeB
            frmProperties.MSFlexGrid1.AddItem "Compartmentalization" & vbTab & .Compartmentalization
            frmProperties.MSFlexGrid1.AddItem "Frame Strength" & vbTab & .FrameStrength
            frmProperties.MSFlexGrid1.AddItem "Materials" & vbTab & .Materials
            frmProperties.MSFlexGrid1.AddItem "Responsive" & vbTab & .Responsive
            frmProperties.MSFlexGrid1.AddItem "Robotic" & vbTab & .Robotic
            frmProperties.MSFlexGrid1.AddItem "Biomechanical" & vbTab & .Biomechanical
            frmProperties.MSFlexGrid1.AddItem "Living Metal" & vbTab & .LivingMetal
            frmProperties.MSFlexGrid1.AddItem "Rotation Space" & vbTab & .RotationSpace
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & Str(.Cost)
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & Str(.Weight)
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & Str(.Volume)
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & Str(.SurfaceArea)
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.Text1.Tag = Key
            
        Case "clsGasbag"
                                   
            frmProperties.MSFlexGrid1.AddItem "Tech Level" & vbTab & .TL
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & Str(.Cost)
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & Str(.Weight)
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & Str(.Volume)
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & Str(.SurfaceArea)
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.Text1.Tag = Key
            
        Case "clsArm"
            frmProperties.MSFlexGrid1.AddItem "Tech Level" & vbTab & .TL
            frmProperties.MSFlexGrid1.AddItem "Frame Strength" & vbTab & .FrameStrength
            frmProperties.MSFlexGrid1.AddItem "Materials" & vbTab & .Materials
            frmProperties.MSFlexGrid1.AddItem "Responsive" & vbTab & .Responsive
            frmProperties.MSFlexGrid1.AddItem "Robotic" & vbTab & .Robotic
            frmProperties.MSFlexGrid1.AddItem "Biomechanical" & vbTab & .Biomechanical
            frmProperties.MSFlexGrid1.AddItem "Living Metal" & vbTab & .LivingMetal
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & Str(.Cost)
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & Str(.Weight)
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & Str(.Volume)
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & Str(.SurfaceArea)
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.Text1.Tag = Key
        '////////////////////////////////////////////
        'Propulsion Systems
        Case "clsWheeledDrivetrain", "clsTrackedDrivetrain", "clsLegDrivetrain", "clsFlexibodyDrivetrain"
            frmProperties.MSFlexGrid1.AddItem "Motive Power" & vbTab & .MotivePower
            frmProperties.MSFlexGrid1.AddItem "Ruggedized" & vbTab & .Ruggedized
            frmProperties.MSFlexGrid1.AddItem "Power Consumption" & vbTab & .PowerReqt
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & .Cost
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & .Weight
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & .Volume
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & .SurfaceArea
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.MSFlexGrid1.AddItem "DR" & vbTab & .DR
            frmProperties.Text1.Tag = Key
            
        Case "clsOrnithopterDrivetrain", "clsTTRRotarDrivetrain", "clsMMRRotarDrivetrain", "clsCARRotarDrivetrain"
            frmProperties.MSFlexGrid1.AddItem "Motive Power" & vbTab & .MotivePower
            frmProperties.MSFlexGrid1.AddItem "Lift" & vbTab & .Lift
            frmProperties.MSFlexGrid1.AddItem "Motive Thrust" & vbTab & .MotiveThrust
            frmProperties.MSFlexGrid1.AddItem "Ruggedized" & vbTab & .Ruggedized
            frmProperties.MSFlexGrid1.AddItem "Power Consumption" & vbTab & .PowerReqt
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & .Cost
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & .Weight
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & .Volume
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & .SurfaceArea
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.MSFlexGrid1.AddItem "DR" & vbTab & .DR
            frmProperties.Text1.Tag = Key

        Case "clsAerialPropeller"
            frmProperties.MSFlexGrid1.AddItem "Motive Power" & vbTab & .MotivePower
            frmProperties.MSFlexGrid1.AddItem "Motive Thrust" & vbTab & .MotiveThrust
            frmProperties.MSFlexGrid1.AddItem "Ruggedized" & vbTab & .Ruggedized
            frmProperties.MSFlexGrid1.AddItem "Power Consumption" & vbTab & .PowerReqt
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & .Cost
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & .Weight
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & .Volume
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & .SurfaceArea
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.MSFlexGrid1.AddItem "DR" & vbTab & .DR
            frmProperties.Text1.Tag = Key
            
        Case "clsDuctedFan"
            frmProperties.MSFlexGrid1.AddItem "Motive Power" & vbTab & .MotivePower
            frmProperties.MSFlexGrid1.AddItem "Motive Thrust" & vbTab & .MotiveThrust
            frmProperties.MSFlexGrid1.AddItem "Lift Engine" & vbTab & .LiftEngine
            frmProperties.MSFlexGrid1.AddItem "Vectored Thurst" & vbTab & .VectoredThrust
            frmProperties.MSFlexGrid1.AddItem "Ruggedized" & vbTab & .Ruggedized
            frmProperties.MSFlexGrid1.AddItem "Power Consumption" & vbTab & .PowerReqt
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & .Cost
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & .Weight
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & .Volume
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & .SurfaceArea
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.MSFlexGrid1.AddItem "DR" & vbTab & .DR
            frmProperties.Text1.Tag = Key
            
        Case "clsAquaticPropulsion"
            frmProperties.MSFlexGrid1.AddItem "Power Requirement" & vbTab & .PowerReqt
            frmProperties.MSFlexGrid1.AddItem "Motive Thrust" & vbTab & .MotiveThrust
            frmProperties.MSFlexGrid1.AddItem "Ruggedized" & vbTab & .Ruggedized
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & .Cost
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & .Weight
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & .Volume
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & .SurfaceArea
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.MSFlexGrid1.AddItem "DR" & vbTab & .DR
            frmProperties.Text1.Tag = Key
            
        Case "clsHarness"
            frmProperties.MSFlexGrid1.AddItem "Motive Power/Thrust" & vbTab & .MotiveThrust
            frmProperties.MSFlexGrid1.AddItem "Combined Beast ST" & vbTab & .BeastST
            frmProperties.MSFlexGrid1.AddItem "Hexes of Animals" & vbTab & .BeastHexes
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & .Cost
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & .Weight
            frmProperties.Text1.Tag = Key
        
        Case "clsMagLevLifter"
            frmProperties.MSFlexGrid1.AddItem "Lift" & vbTab & .Lift
            frmProperties.MSFlexGrid1.AddItem "Ruggedized" & vbTab & .Ruggedized
            frmProperties.MSFlexGrid1.AddItem "Power Consumption" & vbTab & .PowerReqt
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & .Cost
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & .Weight
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & .Volume
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & .SurfaceArea
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.MSFlexGrid1.AddItem "DR" & vbTab & .DR
            frmProperties.Text1.Tag = Key
            
        Case "clsJetEngine"
            frmProperties.MSFlexGrid1.AddItem "Motive Thrust" & vbTab & .MotiveThrust
            frmProperties.MSFlexGrid1.AddItem "Afterburner" & vbTab & .Afterburner
            frmProperties.MSFlexGrid1.AddItem "Lift Engine" & vbTab & .LiftEngine
            frmProperties.MSFlexGrid1.AddItem "Vectored Thurst" & vbTab & .VectoredThrust
            frmProperties.MSFlexGrid1.AddItem "Ruggedized" & vbTab & .Ruggedized
            frmProperties.MSFlexGrid1.AddItem "Fuel Consumption" & vbTab & .PowerReqt
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & .Cost
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & .Weight
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & .Volume
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & .SurfaceArea
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.MSFlexGrid1.AddItem "DR" & vbTab & .DR
            frmProperties.Text1.Tag = Key
            
        Case "clsReactionlessThruster"
            frmProperties.MSFlexGrid1.AddItem "Motive Thrust" & vbTab & .MotiveThrust
            frmProperties.MSFlexGrid1.AddItem "Lift Engine" & vbTab & .LiftEngine
            frmProperties.MSFlexGrid1.AddItem "Vectored Thurst" & vbTab & .VectoredThrust
            frmProperties.MSFlexGrid1.AddItem "Ruggedized" & vbTab & .Ruggedized
            frmProperties.MSFlexGrid1.AddItem "Power Consumption" & vbTab & .PowerReqt
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & .Cost
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & .Weight
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & .Volume
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & .SurfaceArea
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.MSFlexGrid1.AddItem "DR" & vbTab & .DR
            frmProperties.Text1.Tag = Key
           
        Case "clsRocketEngine"
            frmProperties.MSFlexGrid1.AddItem "Motive Thrust" & vbTab & .MotiveThrust
            frmProperties.MSFlexGrid1.AddItem "Lift Engine" & vbTab & .LiftEngine
            frmProperties.MSFlexGrid1.AddItem "Vectored Thurst" & vbTab & .VectoredThrust
            frmProperties.MSFlexGrid1.AddItem "Ruggedized" & vbTab & .Ruggedized
            frmProperties.MSFlexGrid1.AddItem "Fuel Consumption" & vbTab & .FuelReqt
            frmProperties.MSFlexGrid1.AddItem "Power Consumption" & vbTab & .PowerReqt
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & .Cost
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & .Weight
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & .Volume
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & .SurfaceArea
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.MSFlexGrid1.AddItem "DR" & vbTab & .DR
            frmProperties.Text1.Tag = Key

        Case "clsRowingPositions"
            frmProperties.MSFlexGrid1.AddItem "Motive Power/Thrust" & vbTab & .MotiveThrust
            frmProperties.MSFlexGrid1.AddItem "Combined Rowers ST" & vbTab & .RowerST
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & .Cost
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & .Weight
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & .Volume
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & .SurfaceArea
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.MSFlexGrid1.AddItem "DR" & vbTab & .DR
            frmProperties.Text1.Tag = Key

        Case "clsSail", "clsAerialSail"
            frmProperties.MSFlexGrid1.AddItem "Motive Thrust" & vbTab & .MotiveThrust
            frmProperties.MSFlexGrid1.AddItem "Sail Material" & vbTab & .Material
            frmProperties.MSFlexGrid1.AddItem "Wind" & vbTab & .Wind
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & .Cost
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & .Weight
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & .Volume
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & .SurfaceArea
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.MSFlexGrid1.AddItem "DR" & vbTab & .DR
            frmProperties.Text1.Tag = Key
                       
        Case "clsLightSail"
            frmProperties.MSFlexGrid1.AddItem "Thrust" & vbTab & .Thrust
            frmProperties.MSFlexGrid1.AddItem "Sail Size in sq. mi." & vbTab & .Size
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & .Cost
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & .Weight
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & .Volume
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & .SurfaceArea
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.MSFlexGrid1.AddItem "DR" & vbTab & .DR
            frmProperties.Text1.Tag = Key

        Case "clsJumpDrive", "clsTeleportationDrive"
            frmProperties.MSFlexGrid1.AddItem "Capacity" & vbTab & .Capacity
            frmProperties.MSFlexGrid1.AddItem "Power Consumption" & vbTab & .PowerReqt
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & .Cost
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & .Weight
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & .Volume
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & .SurfaceArea
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.MSFlexGrid1.AddItem "DR" & vbTab & .DR
            frmProperties.Text1.Tag = Key
            
        Case "clsHyperDrive"
            frmProperties.MSFlexGrid1.AddItem "Capacity" & vbTab & .Capacity
            frmProperties.MSFlexGrid1.AddItem "Power Conspt. (initial)" & vbTab & .PowerReqt
            frmProperties.MSFlexGrid1.AddItem "Power Conspt.(sustained)" & vbTab & .SustainedPowerReqt
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & .Cost
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & .Weight
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & .Volume
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & .SurfaceArea
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.MSFlexGrid1.AddItem "DR" & vbTab & .DR
            frmProperties.Text1.Tag = Key
        
        Case "clsWarpDrive"
            frmProperties.MSFlexGrid1.AddItem "Warp Thrust Factor" & vbTab & .WarpThrustFactor
            frmProperties.MSFlexGrid1.AddItem "Power Consumption" & vbTab & .PowerReqt
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & .Cost
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & .Weight
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & .Volume
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & .SurfaceArea
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.MSFlexGrid1.AddItem "DR" & vbTab & .DR
            frmProperties.Text1.Tag = Key
            
        Case "clsOrionEngine"
            frmProperties.MSFlexGrid1.AddItem "Pulse Rate" & vbTab & .PulseRate
            frmProperties.MSFlexGrid1.AddItem "Bomb Size" & vbTab & .BombSize
            frmProperties.MSFlexGrid1.AddItem "Lift Engine" & vbTab & .LiftEngine
            frmProperties.MSFlexGrid1.AddItem "Max Thrust Time" & vbTab & .ThrustTime
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & .Cost
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & .Weight
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & .Volume
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & .SurfaceArea
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.MSFlexGrid1.AddItem "DR" & vbTab & .DR
            frmProperties.Text1.Tag = Key
            
        Case "clsParachronicConveyer"
            frmProperties.MSFlexGrid1.AddItem "Conveyor Sub-Type" & vbTab & .SubType
            frmProperties.MSFlexGrid1.AddItem "Power Consumption" & vbTab & .PowerReqt
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & .Cost
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & .Weight
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & .Volume
            frmProperties.MSFlexGrid1.AddItem "Surface Area" & vbTab & .SurfaceArea
            frmProperties.MSFlexGrid1.AddItem "Hit Points" & vbTab & .HitPoints
            frmProperties.MSFlexGrid1.AddItem "DR" & vbTab & .DR
            frmProperties.Text1.Tag = Key
    '/////////////////////////////////////////
    ' Aerostatic Lift Systems
        Case "clsLiftingGas"
            frmProperties.MSFlexGrid1.AddItem "Useful Lift" & vbTab & .Lift
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & .Cost
            frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & .Volume
            frmProperties.Text1.Tag = Key
        
        Case "clsContraGravGenerator"
            frmProperties.MSFlexGrid1.AddItem "Lift" & vbTab & .Lift
            frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & .Cost
            frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & .Weight
            frmProperties.Text1.Tag = Key
    '///////////////////////////////////////////
    'Armor
        
        Case "clsArmor"
            Select Case ComponentsParent
                Case "clsOpenMount"
                    frmProperties.MSFlexGrid1.AddItem "Armor Setting" & vbTab & "Gun Shield"
                    frmProperties.MSFlexGrid1.AddItem "DR" & vbTab & .DR
                    frmProperties.MSFlexGrid1.AddItem "PR" & vbTab & .PR
                    frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & .Cost
                    frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & .Weight
                    frmProperties.MSFlexGrid1.AddItem "Type" & vbTab & .Material
                    frmProperties.MSFlexGrid1.AddItem "Quality" & vbTab & .Quality
                    ArmorOption = OpenFrame
                Case "clsGasbag"
                    frmProperties.MSFlexGrid1.AddItem "DR" & vbTab & .DR
                    frmProperties.MSFlexGrid1.AddItem "PR" & vbTab & .PR
                    frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & .Cost
                    frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & .Weight
                    frmProperties.MSFlexGrid1.AddItem "Type" & vbTab & .Material
                    frmProperties.MSFlexGrid1.AddItem "Quality" & vbTab & .Quality
                    ArmorOption = Location
                Case "clsBody", "clsSuperStructure", "clsTurret", "clsPopTurret", "clsSkid", "clsWheel", "clsTrack", "clsLeg", "clsWing", "clsRotar", "clsHydrofoil", "clsHovercraft", "clsMast", "clsPod", "clsArm"
                    frmProperties.MSFlexGrid1.AddItem "Settings"
                    frmProperties.MSFlexGrid1.CellFontBold = True
                    frmProperties.MSFlexGrid1.AddItem "Armor Setting" & vbTab & .ArmorOption
                    If colVehicle.Item(Key).ArmorOption = "Location" Then
                        frmProperties.MSFlexGrid1.AddItem "DR" & vbTab & .FDR
                        frmProperties.MSFlexGrid1.AddItem "Material" & vbTab & .Material
                        frmProperties.MSFlexGrid1.AddItem "Quality" & vbTab & .Quality
                    Else
                        frmProperties.MSFlexGrid1.AddItem "DR" & vbTab & "(custom)"
                        frmProperties.MSFlexGrid1.AddItem "Material" & vbTab & "(custom)"
                        frmProperties.MSFlexGrid1.AddItem "Quality" & vbTab & "(custom)"
                    End If
                    frmProperties.MSFlexGrid1.AddItem "Radiation Shielding" & vbTab & .Radiation
                    frmProperties.MSFlexGrid1.AddItem "Surface Coating" & vbTab & .Coating
                    frmProperties.MSFlexGrid1.AddItem "Thermal Superconductor" & vbTab & .Thermal
                    frmProperties.MSFlexGrid1.AddItem "Reactive Armor Plating" & vbTab & .RAP
                    frmProperties.MSFlexGrid1.AddItem "Electrified" & vbTab & .Electrified
                    frmProperties.MSFlexGrid1.AddItem "Statistics"
                    frmProperties.MSFlexGrid1.Row = 10
                    frmProperties.MSFlexGrid1.CellFontBold = True
                    If colVehicle.Item(Key).ArmorOption = "Location" Then
                            frmProperties.MSFlexGrid1.AddItem "PR" & vbTab & .FPD
                        Else
                        If component = "clsBody" Then
                        ' Average = FDR / 6
                            frmProperties.MSFlexGrid1.AddItem "Average DR" & vbTab & .FDR / 6
                            frmProperties.MSFlexGrid1.AddItem "Average PD" & vbTab & .FPD / 6
                        Else
                        ' Average = FDR / 5
                            frmProperties.MSFlexGrid1.AddItem "Average DR" & vbTab & .FDR / 5
                            frmProperties.MSFlexGrid1.AddItem "Average PD" & vbTab & .FPD / 5
                        End If
                    End If
                    frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & .Cost
                    frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & .Weight
                    frmProperties.MSFlexGrid1.Row = 1
                    frmProperties.Text1.Tag = Key
                Case Else
                    
                    frmProperties.MSFlexGrid1.AddItem "Average DR" & vbTab & .DR
                    frmProperties.MSFlexGrid1.AddItem "PR" & vbTab & .PD
                    frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & .Cost
                    frmProperties.MSFlexGrid1.AddItem "Weight" & vbTab & .Weight
                    frmProperties.MSFlexGrid1.AddItem "Type" & vbTab & .Material
                    frmProperties.MSFlexGrid1.AddItem "Quality" & vbTab & .Quality
                    frmProperties.Text1.Tag = Key
                    ArmorOption = Location
                End Select
            '/////////////////////////////////////////////
            'Cargo Space
            '/////////////////////////////////////////////
            Case "clsCargo"
                frmProperties.MSFlexGrid1.AddItem "SubType DR" & vbTab & .SubType
                frmProperties.MSFlexGrid1.AddItem "Cargo Space" & vbTab & .CargoSpace
                frmProperties.MSFlexGrid1.AddItem "Volume" & vbTab & .Volume
                If colVehicle.Item(Key).SubType <> "Hidden" Then
                Else
                    frmProperties.MSFlexGrid1.AddItem "Cost" & vbTab & .Cost
                End If
                frmProperties.Text1.Tag = Key
            '//////////////////////////////////////////////
            'Manned Vehicle Components
            '//////////////////////////////////////////////
            
    End Select
End With
End Sub


Public Sub Performance_Update()

Dim i As Integer
Dim iButton As Integer

' clear the grid and set the # of rows in the grid to 0
    frmPerformance.MSFlexGrid1.Clear
    frmPerformance.MSFlexGrid1.Rows = 0
    ' Set the width of the first column
    frmPerformance.MSFlexGrid1.ColWidth(0) = 1900
    ' set the width of the second column so that it fills up the rest of the grids width
    frmPerformance.MSFlexGrid1.ColWidth(1) = frmPerformance.MSFlexGrid1.Width - frmPerformance.MSFlexGrid1.ColWidth(0) - 60
'make sure a profile is selected in the combobox
If frmPerformance.Combo1.Text = "" Then Exit Sub
With colVehicle.Item(frmPerformance.Combo1.Text)

Select Case TypeName(colVehicle.Item(.PerformanceType))
    Case ""
        frmPerformance.Caption = " Air Performance"
        frmPerformance.MSFlexGrid1.AddItem "Top Speed" & vbTab & .aTopSpeed
        frmPerformance.MSFlexGrid1.AddItem "Acceleration" & vbTab & .aAcceleration
        frmPerformance.MSFlexGrid1.AddItem "Deceleration" & vbTab & .aDeceleration
        frmPerformance.MSFlexGrid1.AddItem "Stability" & vbTab & .aStability
        frmPerformance.MSFlexGrid1.AddItem "Maneuverability" & vbTab & .aManeuverability
        frmPerformance.MSFlexGrid1.AddItem "Drag" & vbTab & .aDrag
        frmPerformance.MSFlexGrid1.AddItem "Stall Speed" & vbTab & .aStallSpeed
        frmPerformance.MSFlexGrid1.AddItem "Motive Thrust" & vbTab & .aMotiveThrust
 
    
    Case "clsWheel", "clsTrack", "clsLeg", "clsFlexibodyDrivetrain"
        
        frmPerformance.Caption = " Ground Performance"
        frmPerformance.MSFlexGrid1.AddItem "Top Speed" & vbTab & .gTopSpeed
        frmPerformance.MSFlexGrid1.AddItem "Acceleration" & vbTab & .gAcceleration
        frmPerformance.MSFlexGrid1.AddItem "Deceleration" & vbTab & .gDeceleration
        frmPerformance.MSFlexGrid1.AddItem "Stability" & vbTab & .gStability
        frmPerformance.MSFlexGrid1.AddItem "Maneuverability" & vbTab & .gManeuverability
        frmPerformance.MSFlexGrid1.AddItem "Pressure" & vbTab & .gPressure
        frmPerformance.MSFlexGrid1.AddItem "Off Road" & vbTab & .gOffRoad
    
    Case "clsHovercraft"
    
        frmPerformance.Caption = " Hovercraft Performance"
        frmPerformance.MSFlexGrid1.AddItem "Top Speed" & vbTab & .hTopSpeed
        frmPerformance.MSFlexGrid1.AddItem "Acceleration" & vbTab & .hAcceleration
        frmPerformance.MSFlexGrid1.AddItem "Deceleration" & vbTab & .hDeceleration
        frmPerformance.MSFlexGrid1.AddItem "Stability" & vbTab & .hStability
        frmPerformance.MSFlexGrid1.AddItem "Maneuverability" & vbTab & .hManeuverability
        
    
    Case "clsMagLevLifter"
        
        frmPerformance.Caption = " Mag-Lev Performance"
        frmPerformance.MSFlexGrid1.AddItem "Top Speed" & vbTab & .mTopSpeed
        frmPerformance.MSFlexGrid1.AddItem "Acceleration" & vbTab & .mAcceleration
        frmPerformance.MSFlexGrid1.AddItem "Deceleration" & vbTab & .mDeceleration
        frmPerformance.MSFlexGrid1.AddItem "Stability" & vbTab & .mStability
        frmPerformance.MSFlexGrid1.AddItem "Maneuverability" & vbTab & .mManeuverability
    Case "Water__"
    
        frmPerformance.Caption = " Water Performance"
        frmPerformance.MSFlexGrid1.AddItem "Top Speed" & vbTab & .wTopSpeed
        frmPerformance.MSFlexGrid1.AddItem "Hydrofoil Speed" & vbTab & .wHydrofoilSpeed
        frmPerformance.MSFlexGrid1.AddItem "Planing Speed" & vbTab & .wPlaningSpeed
        frmPerformance.MSFlexGrid1.AddItem "Acceleration" & vbTab & .wAcceleration
        frmPerformance.MSFlexGrid1.AddItem "Deceleration" & vbTab & .wDeceleration
        frmPerformance.MSFlexGrid1.AddItem "Increased Deceleration" & vbTab * .wIDeceleration
        frmPerformance.MSFlexGrid1.AddItem "Stability" & vbTab & .wStability
        frmPerformance.MSFlexGrid1.AddItem "Maneuverability" & vbTab & .wManeuverability
        frmPerformance.MSFlexGrid1.AddItem "Draft" & vbTab & .wDraft
    
    Case "Submerged__"
    
        frmPerformance.Caption = " Submerged Performance"
        frmPerformance.MSFlexGrid1.AddItem "Top Speed" & vbTab & .suTopSpeed
        frmPerformance.MSFlexGrid1.AddItem "Acceleration" & vbTab & .suAcceleration
        frmPerformance.MSFlexGrid1.AddItem "Deceleration" & vbTab & .suDeceleration
        frmPerformance.MSFlexGrid1.AddItem "Stability" & vbTab & .suStability
        frmPerformance.MSFlexGrid1.AddItem "Maneuverability" & vbTab & .suManeuverability
        frmPerformance.MSFlexGrid1.AddItem "Hydro. Drag" & vbTab & .suHydroDrag
        frmPerformance.MSFlexGrid1.AddItem "Crush Depth" & vbTab & .suCrushDepth
         
   Case "Space__"
    
        frmPerformance.Caption = " Space Performance"
        frmPerformance.MSFlexGrid1.AddItem "Top Speed" & vbTab & .sTopSpeed
        frmPerformance.MSFlexGrid1.AddItem "Acceleration" & vbTab & .sAcceleration
        frmPerformance.MSFlexGrid1.AddItem "Deceleration" & vbTab & .sDeceleration
        frmPerformance.MSFlexGrid1.AddItem "Stability" & vbTab & .sStability
        frmPerformance.MSFlexGrid1.AddItem "Maneuverability" & vbTab & .sManeuverability
        frmPerformance.MSFlexGrid1.AddItem "FTL Speed" & vbTab & .sFTLSpeed
End Select
End With
End Sub


Sub UpdatePerformanceListbox()

End Sub

Sub CalcPerformance()
' reads the selected Combo1 text and then determines the Key and Subassembly type for that item
' once the key is determined, it adds the total motive powers and finds the top speed


' it picks the correct loaded weight depending on whether hardpoints is checked or not


End Sub
