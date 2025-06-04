VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum xopErrors
    errMSXMLerror = vbObjectError + 1
    errNoPeristentData
    errUnknownType
    errNoMultiDimArraysSupported
End Enum

Private m_XMLDom40 As DOMDocument40

' IMPORTANT: when serializing objects, we add them to the m_objectCache.  If an object we want to serialize is in the
' cache, we skip it since we dont want to serialize it twice.  Depending on how our program is designed
' we could have each child component referencing its parent.  That would result in parents getting serialized
' every time a child was serialized.  Clearly we dont want this.
Private m_objectCache As Collection '<-- why is this here if its not being used?  Something is wrong i think?
Public KeyManager As clsKeyManager
Public m_sFormatString As String
Private WithEvents m_os As PersistenceManager.ObjectStore
Attribute m_os.VB_VarHelpID = -1


Public Property Let FormatString(ByVal vdata As String)
    p_sFormat = vdata
End Property
Function GetOverDrive() As String
    GetOverDrive = TypeName(Me)
End Function
Public Sub SetMessageTextBox(ByRef vdata As Object)
    Set InfoBox = vdata
End Sub

' ----------------------------------------------------
' FUNCTION: CreateComponent()
'
' PURPOSE:  here we need to load in the XML src
'           Access the Persistance Interface of the object (cIPersist)
'           Restore its properties from the DEF file.
'           if all this is successful, we return happily with an object that is already loaded with its DEF file.
'           The calling function can now continue and load the actual USER saved attributes from the .CMP file
Function CreateComponent(ByRef sClassName As String, ByRef sDefPath As String, ByRef sDefID As String) As cINode
    On Error GoTo error
    Dim oNode As Vehicles.cINode
    Dim lngPtr As Long
    Dim oPersistantObj As cIPersist
    
    'todo: note the sDefPath could be a path INSIDE a PAK file.  Its not necessarily a windows file path.  Thoughts?
    '      For now we wont worry about it. That will be one of the last things we implement
    
    'todo: ulimately, arrange this list so that the most frequently used cases are at top and the ones that are only used
    ' rarely, or even once are at the bottom
    Select Case sClassName
            
        ' -----------------------------------------------------------------------------------------------
        ' -------------------------- NOTE: The following group of components do not support cINode and so
        '                            they are not instanced here, instead they fall thru back to the .RequestObject()
        '                            functin and uses standard old CreateObject()
        '                            IMPORTANT!  Since these are child objects with no Tree nodes dedicated specifically
        '                            for them, they DO NOT and in fact CANNOT use .XML DEF files.  This means
        '                            that they can only act as data stores, which is fine.  To access these child objects
        '                            The parent object's XML file must contain any property items used to access them.
        
        Case CLASSNAME_PROPERTY_ITEM
        Case CLASSNAME_AUTHOR
        
        Case CLASSNAME_STATS
        Case CLASSNAME_DESCRIPTION
        
        ' these following dont implement cInode yet,but they should i think since they are represented with their own nodes
        Case CLASSNAME_SURFACE
            Set oNode = New Vehicles.cSurface
        
        Case CLASSNAME_FEATURE
            Set oNode = New Vehicles.cFeature
            
        Case CLASSNAME_OPTIONS
            Set oNode = New Vehicles.cOptions 'temporarily this does implement cINode but im thinking eventually this class is not needed
        Case CLASSNAME_CREW
            Set oNode = New Vehicles.cCrew
        
        ' -------------------------------------------------------------------------------------------------
        ' ---------------------------  The following group does implement cINode and are instanced here
        'todo: i "should" be able to ditch this entire Select Case and use CreateObject(classname)
        '      cuz otherwise this is just way too much crap
        Case CLASSNAME_VEHICLE
            Set oNode = New Vehicles.cVehicle
        Case CLASSNAME_GROUP
            Set oNode = New Vehicles.cGroup
            
        
        Case CLASSNAME_BODY ' i think classID is coming in as a 0 at the moment
            Set oNode = New Vehicles.aBody
            
        Case CLASSNAME_HAMMOCK
            Set oNode = New Vehicles.aTest
            
        Case "Vehicles.cArmor"
            Set oNode = New Vehicles.cArmor
        Case "Vehicles.cArmorFace"
            Set oNode = New Vehicles.cArmorFace
        Case "Vehicles.cArmorLayer"
            Set oNode = New Vehicles.cArmorLayer
                        
        Case Else
            Debug.Print "clsFactory:CreateComponent() -- Could not find classname '" & sClassName & "' "
            
    End Select
    
    
    If Not oNode Is Nothing Then
        lngPtr = ObjPtr(oNode)
        Debug.Print "clsFactory:CreateComponent() -- Successfully instanced object '" & sClassName & "'  Handle = " & lngPtr & ".  Attempting to load DEFINITION..."
        oNode.Handle = lngPtr
        Debug.Assert oNode.Handle > 0 ' make sure a local variable actually stored the handle and that the interface contained an implementation
        
        ' load the definitions
        Dim oXNode As MSXML2.IXMLDOMNode
        Dim oPersistantObject As cIPersist
        
        Set m_os = New PersistenceManager.ObjectStore  ' this really isnt even necessary as long as we are NOT trying to load an OBJECT from the Def
        Set oPersistantObject = oNode
             
        Debug.Print "clsFactory:CreateComponet() -- " & App.Path & sDefPath
        
        If Not oPersistantObject Is Nothing Then
        ' todo: I can simplify this by using an INLINE command in the component such that
        ' the definition XML content is included in the single stream.  Thats how X3d file format
        ' does things.  You can still reference the file but when the parser gets to it, it loads
        ' the referenced file automatically.
            Call m_os.Deserialize(App.Path & sDefPath, 0, XML_NODE_OBJECT, oPersistantObject)
        End If
        
        Set m_os = Nothing
        Set oPersistantObject = Nothing
        
        ' todo: verify the classname and guid in the DEF match with the args passed in from the CMP
        
        
        ' our component is now instanced and loaded with its definitions
        Set CreateComponent = oNode
    Else
        Debug.Print "clsFactory:CreateComponent() -- Could not instance classname '" & sClassName & "' "
    End If
    Exit Function
error:
        ' if the nodes parent property isnt set, that means its a root node so skip
        ' it and resume next
        
     If err.Number = 20 Then
        Debug.Print "clsFactory.CreateComponent - ERROR #" & err.Number & " " & err.Description & " <0K -- ROOT NODE EXCEPTION>"
     ElseIf err.Number = 457 Then
        ' the items key is not unique. This needs to be trapped in case the user
        'tries to add another performance profile that has the same name
        CreateComponent = False
        Exit Function
     ElseIf err.Number = 13 Then
        Debug.Print "clsFactory.CreateComponent - ERROR #" & err.Number & " " & err.Description & " <Possibly? Class does not implement required interface 'cINode'>"
        ' we want to resume here
     ElseIf err.Number = 0 Then
        Exit Function
     Else
        Debug.Print "clsFactory.CreateComponent - ERROR #" & err.Number & " " & err.Description & " <OK -- .INIT METHOD NOT DEFINED IN ALL OBJECT INTERFACES>"
        'note: some classes dont have the .Init method so they barf, but this is ok.
        'i think its a waste to add the .Init to subs that dont need them
        'just to avoid this error which pretty much only occurs during new vehicle creation (e.g. startup)...
        'todo: actually i dont think this debug.print comment is accurate anymore for the new clsFactory.  There is no .init method to speak of
     End If
     Resume Next
End Function

Private Function IsComponent(ByRef sKey As String) As Long
     '++++++++++++++++++++++++++++
    '+ Here is where we should check for node compatibliity.  Depending on whether we rewrite the way locationcheck is performed,
    '+ this code bit may become obsolete, but for now we need it.
    '+
    '+ Essentially, ONLY component nodes can be added to other Component nodes.  Thats basically it.  And since
    '+ no keys in the tree are repeated, if a component with a certain key value doesnt exist in the Veh.Components collection, then
    '+ its not a component and cannot accept other nodes being dropped on it
    'todo: this will be obsolete soon? already?  we check for node compatibility on the fly when a user drags onto
    ' parent the parent must give permission
    
    Dim sTest As String
    
    On Error GoTo err
    
    If Val(sKey) > 0 Then
        IsComponent = True
    Else
        IsComponent = False
    End If
    
    'sTest = TypeName(Vehicle.Components(sKey))
    'IsComponent = True
    Exit Function
err:
    IsComponent = False
    InfoPrint 1, "Components can only be added to valid component nodes."
End Function

Public Function CreateWeaponLink(ByVal sKey As String, ByRef sName As String)
    Dim oWeaponLink As clsWeaponLink
    
    Set oWeaponLink = New clsWeaponLink
    oWeaponLink.Key = sKey
    oWeaponLink.Description = sName
    
    WeaponProfiles.Add oWeaponLink, sKey
    Set oWeaponLink = Nothing
    
    addweaponlink = True
Exit Function

err:
    Debug.Print "clsFactory:CreateWeaponLink() -- ERROR #" & err.Number & "  " & err.Description
End Function
Public Function CreatePerformanceProfile(ByVal Datatype As Integer, ByVal sKey As String, ByVal sParentKey As String, ByVal nImage As Integer, ByVal sNodeText As String) As Long

    On Error GoTo err
    
    Select Case Datatype
        '//////////////////////////////////////////////////////////////////
    'PERFORMANCE PROFILES
    'Case PerformanceProfile
    '    Set objPerformance = New clsPerformance
    '    Profiles.Add objPerformance, sKey
    '    Set objPerformance = Nothing
        
    Case PERFORMANCEAIR
        Set objPerformanceAir = New clsPerformanceAir
        PerformanceProfiles.Add objPerformanceAir, sKey
        Set objPerformanceAir = Nothing
        
    Case PERFORMANCEHOVER
        Set objPerformanceHover = New clsPerformanceHover
        PerformanceProfiles.Add objPerformanceHover, sKey
        Set objPerformanceHover = Nothing
        
    Case PERFORMANCEMAGLEV
        Set objPerformanceMagLev = New clsPerformanceMagLev
        PerformanceProfiles.Add objPerformanceMagLev, sKey
        Set objPerformanceMagLev = Nothing
        
    Case PERFORMANCEWATER
        Set objPerformanceWater = New clsPerformanceWater
        PerformanceProfiles.Add objPerformanceWater, sKey
        Set objPerformanceWater = Nothing
        
    Case PERFORMANCESUB
        Set objPerformanceSub = New clsPerformanceSubmerged
        PerformanceProfiles.Add objPerformanceSub, sKey
        Set objPerformanceSub = Nothing
        
    Case PERFORMANCELEG
        Set objPerformanceLeg = New clsPerformanceLeg
        PerformanceProfiles.Add objPerformanceLeg, sKey
        Set objPerformanceLeg = Nothing
        
    Case PERFORMANCETRACK
        Set objPerformanceTrack = New clsPerformanceTrack
        PerformanceProfiles.Add objPerformanceTrack, sKey
        Set objPerformanceTrack = Nothing
        
    Case PERFORMANCEWHEEL
        'Set objPerformanceWheel = New clsPerformanceWheel
        Set objPerformanceWheel = CreateObject("Vehicles.clsPerformanceWheel")
        PerformanceProfiles.Add objPerformanceWheel, sKey
        Set objPerformanceWheel = Nothing
        
    Case PERFORMANCEFLEX
        Set objPerformanceFlex = New clsPerformanceFlex
        PerformanceProfiles.Add objPerformanceFlex, sKey
        Set objPerformanceFlex = Nothing
        
    Case PERFORMANCESKID
        Set objPerformanceSkid = New clsPerformanceSkid
        PerformanceProfiles.Add objPerformanceSkid, sKey
        Set objPerformanceSkid = Nothing
        
    Case PERFORMANCESPACE
        Set objPerformanceSpace = New clsPerformanceSpace
        PerformanceProfiles.Add objPerformanceSpace, sKey
        Set objPerformanceSpace = Nothing

    End Select

    'add the property values to the objects
    With PerformanceProfiles.Item(sKey)
        .Key = sKey
        .Parent = sParentKey
        .Datatype = Datatype
        .Description = sNodeText
    End With
    
    AddPerformanceProfile = True
    
    'KeyManager.AddPerformanceProfileKey sKey ' MPJ 07/09/02 OBSOLETE - Performance Profiles are now in a seperate "Performance collection" object and dont need to be tracked via keys for efficiency when taking them out of components collection
    Exit Function
err:
    Debug.Print "clsFactory:CreatePerformanceProfile - ERROR #" & err.Number & " " & err.Description
    
    
End Function

Private Sub Class_Initialize()
    Set KeyManager = New clsKeyManager
End Sub

Private Sub Class_Terminate()
     Set KeyManager = Nothing
End Sub

Private Sub m_os_RequestObject(ByVal Classname As String, ByVal DefPath As String, ByVal Def_ID As String, newObject As PersistenceManager.cIPersist)
    Set newObject = CreateComponent(Classname, DefPath, Def_ID)
End Sub


'Private Function LoadComponentFile(sFilePath As String, ByVal lpTarget As Long) As Boolean
'    Dim oXNode As Object
'    Dim oXSibling As Object
'    Dim oXML As cXML
'    Dim lp As Long
'    Dim uCmp As udtComponent
'
'    lp = lpTarget
'
'    Set oXML = New cXML
'
'    If oXML.Initialize(pavV40) Then
'        If oXML.OpenFromFile(sFilePath, True) Then
'            Set oXNode = oXML.GetRootNode
'            Set oXNode = oXML.GetChildByName("Component", oXNode)
'
'            uCmp.ComponentPath = sFilePath
'            If ReadComponentXML(uCmp, oXNode, oXML) Then
'
'                lp = AddNewComponent(lp, uCmp, oXNode) ' get object instance from factory, add it to target object, add tree node
'
'                ' if the root node got added, restore its children (if any)
'                If lp Then
'                    RestoreChildren oXML, oXNode, lp
'                    Set oXSibling = oXNode.nextSibling
'
'                    ' repeat process for all siblings
'                    Do While Not oXSibling Is Nothing
'                        If oXSibling.nodeName = "Component" Then
'                            If ReadComponentXML(uCmp, oXSibling, oXML) Then
'
'                                lp = AddNewComponent(lp, uCmp, oXSibling)
'                                ' if this sibling has children, restore them too
'                                RestoreChildren oXML, oXSibling, lp
'                            End If
'                        End If
'                        Set oXSibling = oXSibling.nextSibling
'                    Loop
'                End If
'            End If
'        End If
'    End If
'
'    Set oXNode = Nothing
'    Set oXSibling = Nothing  'todo: i dont even need sibling, after this functin works, lets try removing and update oxnode to its own sib
'    Set oXML = Nothing
'
'    LoadComponentFile = True
'End Function
'
'Private Function RestoreChildren(oXML As Object, oNode As Object, ByVal lp As Long) As Object   '  <-- returns parent node
'    Dim oChild As Object
'    Dim lParent As Long
'    Dim uCmp As udtComponent
'    Dim i As Long
'
'    lParent = lp
'
'    For i = 0 To oNode.childNodes.Length - 1
'        Set oChild = oNode.childNodes.Item(i)
'        If oChild.nodeName = "Component" Then
'
'            If ReadComponentXML(uCmp, oChild, oXML) Then
'                lParent = AddNewComponent(lParent, uCmp, oChild)
'                If oChild.childNodes.Length >= 1 Then
'                    RestoreChildren oXML, oChild, lParent
'                End If
'            End If
'        End If
'    Next
'    Set oChild = Nothing
'End Function

'Private Function AddNewComponent(ByRef lpParent As Long, uCmp As udtComponent, oXNode As Object) As Long
'    Dim sKey As String
'    Dim sParentKey As String
'    Dim sNodeText As String
'    'Dim sLocation As String
'    'Dim sSavedComponentPath As String
'    Dim lngImageIndex As Long
'    Dim oComponent As Vehicles.cIComponent
'    Dim oParentComponent As Vehicles.cIComponent
'    Dim lngPtr As Long
'    Const LNG_LENGTH = 4
'    Dim sClassName As String
'
'
'    If Not NodeCountExceeded Then
'        sClassName = uCmp.Classname
'        sParentKey = KeyFromLong(lpParent)
'        ' attempting to attach to a valid parent node type?
'        If IsComponent(sParentKey) Then
'
'            CopyMemory oParentComponent, lpParent, LNG_LENGTH
'
'            ' retrieve from factory, the type of object we want to add <-- needs to handle loopage.  So we need
'            ' to open and extract the ClassID's ourselves actually...
'            Set oComponent = Vehicle.Factory.CreateComponent(sClassName)
'            If Not oComponent Is Nothing Then
'                If oComponent.Load(oXNode) Then
'
'                    If oParentComponent.addChild(oComponent) Then    ' need to restore XML before addChild right?  XML might hold rules?
'                        lngPtr = ObjPtr(oComponent)
'                        sKey = KeyFromLong(lngPtr)
'                        sNodeText = uCmp.Text  'ListView1.SelectedItem
'                        lngImageIndex = ImageList1.ListImages(uCmp.IconPath).index
'                        AddNewChildNode sNodeText, lngImageIndex, sKey, sParentKey, 0
'
'
'                        ' Make the newly dragged over item the selected node
'                        treeVehicle.Nodes.Item(sKey).Selected = True
'                        treeVehicle.DropHighlight = treeVehicle.Nodes(sKey)
'                        Call SetActiveNode
'                        treeVehicle.SelectedItem.Expanded = True
'                        AddNewComponent = lngPtr
'                    End If
'                End If
'            End If
'            CopyMemory oParentComponent, 0&, LNG_LENGTH
'        End If
'    End If
'     mbIndrag = False
'    Exit Function
'err:
'    mbIndrag = False
'
'End Function

'    Case Body
'        Set objBody = New clsBody
'        Components.Add objBody, CurrentItem
'        Set objBody = Nothing
            
''''    '////////////////////////////////////////////////////////////////
''''    Case BattlesuitSystem
''''        Set objBattleSuitsystem = New clsBattlesuitSystem
''''        Components.Add objBattleSuitsystem, CurrentItem
''''        Set objBattleSuitsystem = Nothing
''''
''''    Case SimpleCustom
''''        Set objSimpleCustom = New clsSimpleCustom
''''        Components.Add objSimpleCustom, CurrentItem
''''        Set objSimpleCustom = Nothing
''''
''''    Case GroupComponent
''''        Set objGroupComponent = New clsGroup
''''        Components.Add objGroupComponent, CurrentItem
''''        Set objGroupComponent = Nothing
''''
''''    '//////////////////////////////////////////////////////////////////
''''    'SubAssemblies
''''    '//////////////////////////////////////////////////////////////////
''''
''''
''''
''''    Case Wheel
''''        Set objWheel = New clsWheel
''''        Components.Add objWheel, CurrentItem  ' add the new object to the collection
''''        Set objWheel = Nothing
''''
''''    Case Skid
''''        Set objSkid = New clsSkid
''''        Components.Add objSkid, CurrentItem  ' add the new object to the collection
''''        Set objSkid = Nothing
''''
''''    Case Track
''''        Set objTrack = New clsTrack
''''        Components.Add objTrack, CurrentItem  ' add the new object to the collection
''''        Set objTrack = Nothing
''''
''''    Case Hydrofoil
''''        Set objHydrofoil = New clsHydrofoil
''''        Components.Add objHydrofoil, CurrentItem  ' add the new object to the collection
''''        Set objHydrofoil = Nothing
''''
''''    Case Hovercraft
''''        Set objHovercraft = New clsHovercraft
''''        Components.Add objHovercraft, CurrentItem  ' add the new object to the collection
''''        Set objHovercraft = Nothing
''''
''''    Case Leg
''''        Set objLeg = New clsLeg
''''        Components.Add objLeg, CurrentItem  ' add the new object to the collection
''''        Set objLeg = Nothing
''''
''''    Case Arm
''''        Set objArm = New clsArm
''''        Components.Add objArm, CurrentItem  ' add the new object to the collection
''''        Set objArm = Nothing
''''
''''    Case AutogyroRotor, TTRotor, CARotor, MMRotor
''''        Set objRotor = New clsRotor
''''        Components.Add objRotor, CurrentItem  ' add the new object to the collection
''''        Set objRotor = Nothing
''''
''''    Case Wing
''''        Set objWing = New clsWing
''''        Components.Add objWing, CurrentItem  ' add the new object to the collection
''''        Set objWing = Nothing
''''
''''    Case Mast
''''        Set objMast = New clsMast
''''        Components.Add objMast, CurrentItem  ' add the new object to the collection
''''        Set objMast = Nothing
''''
''''    Case Superstructure
''''        Set objSuperStructure = New clsSuperStructure
''''        Components.Add objSuperStructure, CurrentItem  ' add the new object to the collection
''''        Set objSuperStructure = Nothing
''''
''''    Case Turret ' Turret
''''        Set objTurret = New clsTurret
''''        Components.Add objTurret, CurrentItem  ' add the new object to the collection
''''        Set objTurret = Nothing
''''
''''    Case Popturret ' Popturret
''''        Set objPopturret = New clsPopTurret
''''        Components.Add objPopturret, CurrentItem  ' add the new object to the collection
''''        Set objPopturret = Nothing
''''
''''    Case OpenMount
''''        Set objOpenMount = New clsOpenMount
''''        Components.Add objOpenMount, CurrentItem  ' add the new object to the collection
''''        Set objOpenMount = Nothing
''''
''''    Case Gasbag
''''        Set objGasbag = New clsGasbag
''''        Components.Add objGasbag, CurrentItem  ' add the new object to the collection
''''        Set objGasbag = Nothing
''''
''''    Case Pod
''''        Set objPod = New clsPod
''''        Components.Add objPod, CurrentItem  ' add the new object to the collection
''''        Set objPod = Nothing
''''
''''    Case equipmentPod
''''        Set objequipmentPod = New clsEquipmentPod
''''        Components.Add objequipmentPod, CurrentItem ' add the new object
''''        Set objequipmentPod = Nothing
''''
''''    Case Cargo
''''        Set objCargo = New clsCargo
''''        Components.Add objCargo, CurrentItem ' add the new object
''''        Set objCargo = Nothing
''''
''''    Case SideCar ' Sidecar
''''        Set objSideCar = New clsSideCar
''''        Components.Add objSideCar, CurrentItem
''''        Set objSideCar = Nothing
''''
''''    Case SolarPanel 'Solar Panel
''''        Set objSolarPanel = New clsSolarPanel
''''        Components.Add objSolarPanel, CurrentItem
''''        Set objSolarPanel = Nothing
''''
''''    Case SolarCellArray
''''        Set objSolarCellArray = New clsSolarCellArray
''''        Components.Add objSolarCellArray, CurrentItem
''''        Set objSolarCellArray = Nothing
''''
''''    '//////////////////////////////////////////////////////////////////
''''    'Armor
''''    '//////////////////////////////////////////////////////////////////
''''    Case ArmorOverall, ArmorComponent, ArmorLocation, ArmorBasicFacing, ArmorComplexFacing, ArmorGunShield, ArmorWheelGuard, ArmorOpenFrame
''''        Set objArmor = New clsArmor
''''        Components.Add objArmor, CurrentItem ' add the new object
''''        Set objArmor = Nothing
''''
''''    '//////////////////////////////////////////////////////////////////
''''    'Weapon and Weapon Accessories
''''    '//////////////////////////////////////////////////////////////////
''''
''''    Case StoneThrower, BoltThrower, RepeatingBoltThrower
''''        Set objStoneBoltThrower = New clsWeaponStoneBoltThrower
''''        Components.Add objStoneBoltThrower, CurrentItem
''''        Set objStoneBoltThrower = Nothing
''''
''''    Case MuzzleLoader, BreechLoader, ManualRepeater, Revolver, MechanicalGatling, SlowAutoloader, FastAutoloader, lightAutomatic, HeavyAutomatic, ElectricGatling
''''        Set objGun = New clsWeaponGun
''''        Components.Add objGun, CurrentItem
''''        Set objGun = Nothing
''''
''''    Case BlueGreenLaser, RainbowLaser, Laser, UVLaser, IRLaser, Disruptor, ChargedParticleBeam, NeutralParticleBeam, _
''''        Flamer, Screamer, Stunner, ParalysisBeam, XRayLaser, FusionBeam, _
''''        GravityBeam, AntiparticleBeam, Graser, Disintegrator, Displacer, BeamedPowerTransmitter, _
''''        MilitaryParalysisBeam
''''        Set objBeamWeapon = New clsWeaponBeam
''''        Components.Add objBeamWeapon, CurrentItem
''''        Set objBeamWeapon = Nothing
''''
''''    Case IronBomb, RetardedBomb, SmartBomb, SelfDestructSystem, ContactMine, ProximityMine, _
''''        PressureTriggerMine, CommandTriggerMine, SmartTriggerMine, _
''''        UnGuidedMissile, UnGuidedTorpedo, GuidedMissile, GuidedTorpedo
''''        Set objTorpMissile = New clsWeaponTorpMissile
''''        Components.Add objTorpMissile, CurrentItem
''''        Set objTorpMissile = Nothing
''''
''''    Case FlameThrower, WaterCannon
''''        Set objLiquidProjector = New clsWeaponLiquidProjector
''''        Components.Add objLiquidProjector, CurrentItem
''''        Set objLiquidProjector = Nothing
''''
''''    Case DisposableLauncher, MuzzleloadingLauncher, BreechloadingLauncher, _
''''        ManualRepeaterLauncher, SlowAutoLoaderLauncher, FastAutoLoaderLauncher, _
''''        RevolverLauncher, lightAutomaticLauncher, HeavyAutomaticLauncher
''''        Set objLauncher = New clsWeaponLauncher
''''        Components.Add objLauncher, CurrentItem
''''        Set objLauncher = Nothing
''''
''''    Case PartialStabilizationGear, FullStabilizationGear
''''        Set objStabilizationGear = New clsStabilizationGear
''''        Components.Add objStabilizationGear, CurrentItem
''''        Set objStabilizationGear = Nothing
''''
''''    Case UniversalMount, CasemateMount, DoorMount, Cyberslave
''''        Set objWeaponMount = New clsWeaponMount
''''        Components.Add objWeaponMount, CurrentItem
''''        Set objWeaponMount = Nothing
''''
''''    Case AntiBlastMagazine
''''        Set objAntiBlastMagazine = New clsAntiBlastMagazine
''''        Components.Add objAntiBlastMagazine, CurrentItem
''''        Set objAntiBlastMagazine = Nothing
''''
''''    Case HardPoint, WeaponBay
''''        Set objHardpoint = New clsHardPoint
''''        Components.Add objHardpoint, CurrentItem
''''        Set objHardpoint = Nothing
''''
''''    Case Ammunition
''''        Set objAmmunition = New clsWeaponAmmunition
''''        Components.Add objAmmunition, CurrentItem
''''        Set objAmmunition = Nothing
''''
''''    Case WeaponLink
''''        Set objWeaponLink = New clsWeaponLink
''''        Components.Add objWeaponLink, CurrentItem
''''        Set objWeaponLink = Nothing
''''
''''    '/////////////////////////////////////////////////////////////////
''''    'Propulsion Systems
''''    '/////////////////////////////////////////////////////////////////
''''    Case WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, FlexibodyDrivetrain, TrackedDrivetrain, LegDrivetrain
''''        Set objGroundDrivetrain = New clsGroundDrivetrain
''''        Components.Add objGroundDrivetrain, CurrentItem
''''        Set objGroundDrivetrain = Nothing
''''
''''    Case CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain
''''        Set objHelicopterDrivetrain = New clsHelicopterDrivetrain
''''        Components.Add objHelicopterDrivetrain, CurrentItem
''''        Set objHelicopterDrivetrain = Nothing
''''
''''    Case OrnithopterDrivetrain
''''        Set objOrnithopterDrivetrain = New clsOrnithopterDrivetrain
''''        Components.Add objOrnithopterDrivetrain, CurrentItem
''''        Set objOrnithopterDrivetrain = Nothing
''''
''''    Case AerialPropeller
''''        Set objAirscrewDrivetrain = New clsAirscrewDrivetrain
''''        Components.Add objAirscrewDrivetrain, CurrentItem
''''        Set objAirscrewDrivetrain = Nothing
''''
''''    Case DuctedFan
''''        Set objDuctedFan = New clsDuctedFan
''''        Components.Add objDuctedFan, CurrentItem
''''        Set objDuctedFan = Nothing
''''
''''    Case PaddleWheel, ScrewPropeller, lightScrewPropeller, DuctedPropeller, Hydrojet, MHDTunnel
''''        Set objAquaticPropulsion = New clsAquaticPropulsion
''''        Components.Add objAquaticPropulsion, CurrentItem
''''        Set objAquaticPropulsion = Nothing
''''
''''    Case RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness
''''        Set objHarness = New clsHarness
''''        Components.Add objHarness, CurrentItem
''''        Set objHarness = Nothing
''''
''''    'Case Animal
''''    '    Set objAnimal = New clsAnimal
''''    '    Components.Add objAnimal, CurrentItem
''''    '    set objanimal =  nothing
''''
''''    Case MagLevLifter
''''        Set objMagLevLifter = New clsMagLevLifter
''''        Components.Add objMagLevLifter, CurrentItem
''''        Set objMagLevLifter = Nothing
''''
''''    Case Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan, FusionAirRam
''''        Set objJetEngine = New clsJetEngine
''''        Components.Add objJetEngine, CurrentItem
''''        Set objJetEngine = Nothing
''''
''''    Case StandardThruster, SuperThruster, MegaThruster
''''        Set objReactionlessThruster = New clsReactionlessThruster
''''        Components.Add objReactionlessThruster, CurrentItem
''''        Set objReactionlessThruster = Nothing
''''
''''    Case LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal, AntimatterPion
''''        Set objRocketEngine = New clsRocketEngine
''''        Components.Add objRocketEngine, CurrentItem
''''        Set objRocketEngine = Nothing
''''
''''    Case RowingPositions
''''        Set objRowingPositions = New clsRowingPositions
''''        Components.Add objRowingPositions, CurrentItem
''''        Set objRowingPositions = Nothing
''''
''''    Case ForeandAftRig, SquareRig, FullRig, AerialSail, AerialSailForeAftRig
''''        Set objSail = New clsSail
''''        Components.Add objSail, CurrentItem
''''        Set objSail = Nothing
''''
''''    Case lightSail
''''        Set objlightSail = New clsLightSail
''''        Components.Add objlightSail, CurrentItem
''''        Set objlightSail = Nothing
''''
''''    Case SolidRocketEngine
''''        Set objSolidRocketEngine = New clsSolidRocketEngine
''''        Components.Add objSolidRocketEngine, CurrentItem
''''        Set objSolidRocketEngine = Nothing
''''
''''    Case OrionEngine
''''        Set objOrionEngine = New clsOrionEngine
''''        Components.Add objOrionEngine, CurrentItem
''''        Set objOrionEngine = Nothing
''''
''''    Case TeleportationDrive, Hyperdrive, JumpDrive, WarpDrive, QuantumConveyor, SubQuantumConveyor, TwoQuantumConveyor
''''        Set objStarDrive = New clsStarDrive
''''        Components.Add objStarDrive, CurrentItem
''''        Set objStarDrive = Nothing
''''
''''    '/////////////////////////////////////////////////////////////////
''''    'Aerostatic Lift Systems
''''    '/////////////////////////////////////////////////////////////////
''''    Case ContraGravGenerator
''''        Set objContraGravGenerator = New clsContraGravGenerator
''''        Components.Add objContraGravGenerator, CurrentItem
''''        Set objContraGravGenerator = Nothing
''''
''''    Case HotAir, Hydrogen, Helium
''''        Set objLiftingGas = New clsLiftingGas
''''        Components.Add objLiftingGas, CurrentItem
''''        Set objLiftingGas = Nothing
''''
''''    '/////////////////////////////////////////////////////////////////
''''    'Instruments and Electronics
''''    '/////////////////////////////////////////////////////////////////
''''    Case RadioDirectionFinder, RadioCommunicator, TightBeamRadio, VLFRadio, CellularPhone, CellularPhonewithRadio, RadioJammer, ElfReceiver, LaserCommunicator, NeutrinoCommunicator, GravityRippleCommunicator
''''        Set objCommunicator = New clsCommunicator
''''        Components.Add objCommunicator, CurrentItem
''''        Set objCommunicator = Nothing
''''
''''    Case Headlight, Searchlight, InfraredSearchlight
''''        Set objSearchlight = New clsSearchlight
''''        Components.Add objSearchlight, CurrentItem
''''        Set objSearchlight = Nothing
''''
''''    Case AstronomicalInstruments, Telescope, lightAmplification, LowlightTV, ExtendableSensorPeriscope
''''        Set objVisualAugmentationSystem = New clsVisualAugmentationSystem
''''        Components.Add objVisualAugmentationSystem, CurrentItem
''''        Set objVisualAugmentationSystem = Nothing
''''
''''    Case Radar, Ladar, NavigationalRadar, AntiCollisionRadar, AESA, LowResImagingRadar, HiResImagingRadar
''''        Set objRadarandLadar = New clsRadarandLadar
''''        Components.Add objRadarandLadar, CurrentItem
''''        Set objRadarandLadar = Nothing
''''
''''    Case ActiveSonar, PassiveSonar
''''        Set objSonar = New clsSonar
''''        Components.Add objSonar, CurrentItem
''''        Set objSonar = Nothing
''''
''''    Case PassiveInfrared, Thermograph, PassiveRadar, PESA
''''        Set objThermPassElectromag = New clsThermPassElectromag
''''        Components.Add objThermPassElectromag, CurrentItem
''''        Set objThermPassElectromag = Nothing
''''
''''    Case Geophone, MAD, MultiScanner, ChemScanner, RadScanner, BioScanner, GravScanner
''''        Set objOtherSensor = New clsOtherSensor
''''        Components.Add objOtherSensor, CurrentItem
''''        Set objOtherSensor = Nothing
''''
''''    Case RangingSoundDetector, SurveillanceSoundDetector
''''        Set objSoundDetector = New clsSoundDetector
''''        Components.Add objSoundDetector, CurrentItem
''''        Set objSoundDetector = Nothing
''''
''''    Case MeteorologicalInstruments, LowResPlanetarySurveyArray, MedResPlanetarySurveyArray, HighResPlanetarySurveyArray
''''        Set objScientificSensor = New clsScientificSensor
''''        Components.Add objScientificSensor, CurrentItem
''''        Set objScientificSensor = Nothing
''''
''''    Case SoundSystem, FlightRecorder, VehicleCamera, DigitalVehicleCamera, ReconCamera, DigitalReconCamera
''''        Set objAudioVisualSystem = New clsAudioVisualSystem
''''        Components.Add objAudioVisualSystem, CurrentItem
''''        Set objAudioVisualSystem = Nothing
''''
''''    Case NavigationInstruments, AutoPilot, IFF, Transponder, INS, GPS, MilitaryGPS, TFR
''''        Set objNavigationSystem = New clsNavigationSystem
''''        Components.Add objNavigationSystem, CurrentItem
''''        Set objNavigationSystem = Nothing
''''
''''    Case ImprovedOpticalBombSight, AdvancedOpticalBombSight, OpticalBombSight, FireDirectionCenter, HUDWAC, PupilHUDWAC, LaserRangeFinder, LaserDesignator, LaserSpotTracker
''''        Set objTargetingSystem = New clsTargetingSystem
''''        Components.Add objTargetingSystem, CurrentItem
''''        Set objTargetingSystem = Nothing
''''
''''    Case RadarDetector, LaserSensor, LaserRadarDetector, AreaRadarJammer, _
''''        DeceptiveRadarJammer, InfraredJammer, ChaffDecoyDischarger, SmokeDecoyDischarger, _
''''        FlareDecoyDischarger, SonarDecoyDischarger, HotSmokeDecoyDischarger, _
''''        PrismDecoyDischarger, BlackOutGasDecoyDischarger, RadarReflector, BlipEnhancer, TEMPEST
''''        Set objElectronicCountermeasure = New clsElectronicCountermeasure
''''        Components.Add objElectronicCountermeasure, CurrentItem
''''        Set objElectronicCountermeasure = Nothing
''''
''''    Case DecoyChaff, DecoySmoke, DecoyFlares, DecoySonarDecoy, DecoyHotSmoke, DecoyPrism, DecoyBlackOutGas
''''        Set objDecoyReload = New clsDecoyReload
''''        Components.Add objDecoyReload, CurrentItem
''''        Set objDecoyReload = Nothing
''''
''''    Case MacroFrame, MainFrame, MicroFrame, MiniComputer, SmallComputer
''''        Set objComputer = New clsComputer
''''        Components.Add objComputer, CurrentItem
''''        Set objComputer = Nothing
''''
''''    Case Terminal
''''        Set objTerminal = New clsTerminal
''''        Components.Add objTerminal, CurrentItem
''''        Set objTerminal = Nothing
''''
''''    Case FireDirectionSoftware, DatabaseSoftware, CartographySoftware, ComputerNavigationSoftware, DatalinkSoftware, TargetingSoftware, TransmissionProfilingSoftware, GunnerSoftware, DamageControlSoftware, PersonalitySimulationSoftwareFull, RobotSkillProgramsPhysical, RoutineVehicleOperationSoftwarePilot, PersonalitySimulationLimited, RoutineVehicleOperationSoftwareOther, RobotSkillProgramsMental, HoloventureProgram
''''        Set objSoftware = New clsSoftware
''''        Components.Add objSoftware, CurrentItem
''''        Set objSoftware = Nothing
''''
''''    Case SurgicalInterface, InterfaceWeb, AutoInterfaceWeb, SocketInterface, NeuralInductionField
''''        Set objNeuralInterfaceSystem = New clsNeuralInterfaceSystem
''''        Components.Add objNeuralInterfaceSystem, CurrentItem
''''        Set objNeuralInterfaceSystem = Nothing
''''
''''    Case DeflectorField, ForceScreen, VariableForceScreen
''''        Set objShields = New clsShields
''''        Components.Add objShields, CurrentItem
''''        Set objShields = Nothing
''''
''''    '/////////////////////////////////////////////////////////////////
''''    'Miscellaneous equipment
''''    '/////////////////////////////////////////////////////////////////
''''    Case ArmMotor
''''    Set objArmMotor = New clsArmMotor
''''        Components.Add objArmMotor, CurrentItem
''''        Set objArmMotor = Nothing
''''
''''    Case FireExtinguisherSystem, FullFireSuppressionSystem, CompactFireSuppressionSystem
''''        Set objFireExtinguisher = New clsFireExtinguisher
''''        Components.Add objFireExtinguisher, CurrentItem
''''        Set objFireExtinguisher = Nothing
''''
''''    Case BilgePump
''''        Set objBilgePump = New clsBilgePump
''''        Components.Add objBilgePump, CurrentItem
''''        Set objBilgePump = Nothing
''''
''''    Case CompleteWorkshop, MechanicWorkshop, EngineeringWorkshop, ElectronicsWorkshop, ArmouryWorkshop, CompleteMiniWorkshop, ScienceLab, MiniMechanicWorkshop, MiniElectronicsWorkshop, MiniEngineeringWorkshop, MiniArmouryWorkshop
''''        Set objLabandWorkshop = New clsLabandWorkshop
''''        Components.Add objLabandWorkshop, CurrentItem
''''        Set objLabandWorkshop = Nothing
''''
''''    Case ExtendableLadder, Crane, Winch, PowerShovel, WreckingCrane, ForkLift, VehicularBridge, LaunchCatapult, SkyHook, Bore, SuperBore, TractorBeam, PressorBeam, CombinationBeam, CraneWithElectroMagnet
''''        Set objHeavyequipment = New clsHeavyEquipment
''''        Components.Add objHeavyequipment, CurrentItem
''''        Set objHeavyequipment = Nothing
''''
''''    Case OperatingRoom, StretcherPallet, EmergencySupportUnit, EmergencylightsandSiren, CryonicCapsule, Automed, DiagnosisTable
''''        Set objEmergencyMedicalequipment = New clsEmergencyMedicalEquipment
''''        Components.Add objEmergencyMedicalequipment, CurrentItem
''''        Set objEmergencyMedicalequipment = Nothing
''''
''''    Case Stage, Hall, BarRoom, ConferenceRoom, MovieScreenandProjector, MovieScreenandProjectorSmall, HoloventureZone
''''        Set objEntertainmentFacility = New clsEntertainmentFacility
''''        Components.Add objEntertainmentFacility, CurrentItem
''''        'door and hatch have been removed from below.  These should now be entered as 'Details' by user in Options dialog
''''        Set objEntertainmentFacility = Nothing
''''
''''    Case CargoRamp, Airlock, MembraneAirlock, Forcelock, PassageTube, ArmoredPassageTube
''''        Set objVehicleAccess = New clsVehicleAccess
''''        Components.Add objVehicleAccess, CurrentItem
''''        Set objVehicleAccess = Nothing
''''
''''    Case TeleportProjector
''''        Set objTeleportProjector = New clsTeleportProjector
''''        Components.Add objTeleportProjector, CurrentItem
''''        Set objTeleportProjector = Nothing
''''
''''    Case BrigsandRestraints, BurglarAlarm, HighSecurityAlarm, MutableLicensePlate, OilSprayer, PaintSprayer, SmokeScreen, SpikeDropper
''''        Set objSecurityDirtyTrick = New clsSecurityDirtyTrick
''''        Components.Add objSecurityDirtyTrick, CurrentItem
''''        Set objSecurityDirtyTrick = Nothing
''''
''''    Case VehicleBay, HangerBay, DryDock, SpaceDock, ExternalCradle
''''        Set objVehicleStorage = New clsVehicleStorage
''''        Components.Add objVehicleStorage, CurrentItem
''''        Set objVehicleStorage = Nothing
''''
''''    Case ArrestorHook, VehicularParachute
''''        Set objLandingAid = New clsLandingAid
''''        Components.Add objLandingAid, CurrentItem
''''        Set objLandingAid = Nothing
''''
''''    Case RefuellingProbe, RefuellingDrogue, FuelElectrolysisSystem, HydrogenFuelScoop, AtmosphereProcessor
''''        Set objFuelAccessory = New clsFuelAccessory
''''        Components.Add objFuelAccessory, CurrentItem
''''        Set objFuelAccessory = Nothing
''''
''''    Case NuclearDamper
''''        Set objNuclearDamper = New clsNuclearDamper
''''        Components.Add objNuclearDamper, CurrentItem
''''        Set objNuclearDamper = Nothing
''''
''''    Case SmallRealityStabilizer, MediumRealityStabilizer, HeavyRealityStabilizer
''''        Set objRealityStabilizer = New clsRealityStabilizer
''''        Components.Add objRealityStabilizer, CurrentItem
''''        Set objRealityStabilizer = Nothing
''''
''''    Case ModularSocket, Module
''''        Set objModularSocket = New clsModularSocket
''''        Components.Add objModularSocket, CurrentItem
''''        Set objModularSocket = Nothing
''''
''''    '/////////////////////////////////////////////////////////////////
''''    'Crew, Passengers, Seating, Seats, Stations
''''    '/////////////////////////////////////////////////////////////////
''''    Case PrimitiveManeuverControl, ElectronicDivingControl, ComputerizedDivingControl, _
''''        MechanicalManeuverControl, ElectronicManeuverControl, _
''''        ComputerizedManeuverControl, MechanicalDivingControl
''''        'duplicate is now a property of the above controls instead of seperate
''''        'DuplicateMechanicalControl, DuplicateElectronicControl, DuplicateComputerizedControl,  DuplicateDivingControl
''''        Set objManeuverControl = New clsManeuverControl
''''        Components.Add objManeuverControl, CurrentItem
''''        Set objManeuverControl = Nothing
''''
''''    Case CrampedCrewStation, NormalCrewStation, RoomyCrewStation, CycleCrewStation, HarnessCrewStation
''''        Set objCrewStation = New clsCrewStation
''''        Components.Add objCrewStation, CurrentItem
''''        Set objCrewStation = Nothing
''''
''''    ' TODO: Temporarily removed Hammock from here so that we can use it for testing
''''    Case CrampedSeat, NormalSeat, RoomySeat, CrampedStandingRoom, NormalStandingRoom, RoomyStandingRoom, CycleSeat, Bunk, Cabin, LuxuryCabin, Suite, LuxurySuite, SmallGalley
''''        Set objAccommodation = New clsAccommodation
''''        Components.Add objAccommodation, CurrentItem
''''        Set objAccommodation = Nothing
''''
''''    Case TotalLifeSystem, ArtificialGravityUnit, EnvironmentalControl, NBCKit, LimitedLifeSystem, FullLifeSystem
''''        Set objEnvironmentalSystem = New clsEnvironmentalSystem
''''        Components.Add objEnvironmentalSystem, CurrentItem
''''        Set objEnvironmentalSystem = Nothing
''''
''''    Case Provisions
''''        Set objProvisions = New clsProvisions
''''        Components.Add objProvisions, CurrentItem
''''        Set objProvisions = Nothing
''''
''''    Case EjectionSeat, CrewEscapeCapsule, Airbag, CrashWeb, WombTank, GravityWeb, GravCompensator
''''        Set objSafetySystem = New clsSafetySystem
''''        Components.Add objSafetySystem, CurrentItem
''''        Set objSafetySystem = Nothing
''''
''''    Case BattlesuitSystem, FormFittingBattleSuitSystem
''''        Set objBattleSuitsystem = New clsBattlesuitSystem
''''        Components.Add objBattleSuitsystem, CurrentItem
''''        Set objBattleSuitsystem = Nothing
''''
''''    '/////////////////////////////////////////////////////////////////
''''    'Power and Fuel Components
''''    '/////////////////////////////////////////////////////////////////
''''    Case MuscleEngine
''''        Set objMuscleEngine = New clsMuscleEngine
''''        Components.Add objMuscleEngine, CurrentItem
''''        Set objMuscleEngine = Nothing
''''
''''    Case GasolineEngine, HPGasolineEngine, TurboGasolineEngine, SuperGasolineEngine, TurboHPGasolineEngine, SuperHPGasolineEngine, StandardDieselEngine, TurboStandardDieselEngine, MarineDieselEngine, HPDieselEngine, TurboHPDieselEngine, CeramicEngine, TurboCeramicEngine, SuperCeramicEngine, HPCeramicEngine, TurboHPCeramicEngine, SuperHPCeramicEngine, HydrogenCombustionEngine
''''        Set objInternalCombustionEngine = New clsInternalCombustionEngine
''''        Components.Add objInternalCombustionEngine, CurrentItem
''''        Set objInternalCombustionEngine = Nothing
''''
''''    Case EarlySteamEngine, ForcedDraftSteamEngine, TripleExpansionSteamEngine, SteamTurbine
''''        Set objSteamEngine = New clsSteamEngine
''''        Components.Add objSteamEngine, CurrentItem
''''        Set objSteamEngine = Nothing
''''
''''    Case StandardGasTurbine, HPGasTurbine, OptimizedGasTurbine, StandardMHDTurbine, HPMHDTurbine
''''        Set objGasandMHDTurbine = New clsGasandMHDTurbine
''''        Components.Add objGasandMHDTurbine, CurrentItem
''''        Set objGasandMHDTurbine = Nothing
''''
''''    Case FuelCell
''''        Set objFuelCell = New clsFuelCell
''''        Components.Add objFuelCell, CurrentItem
''''        Set objFuelCell = Nothing
''''
''''    Case FissionReactor, RTGReactor, NPU, FusionReactor, AntimatterReactor, TotalConversionPowerPlant, CosmicPowerPlant
''''        Set objReactor = New clsReactor
''''        Components.Add objReactor, CurrentItem
''''        Set objReactor = Nothing
''''
''''    Case Soulburner, ElementalFurnace, ManaEngine, Carnivore, Herbivore, Omnivore, Vampire
''''        Set objExoticPowerPlant = New clsExoticPowerPlant
''''        Components.Add objExoticPowerPlant, CurrentItem
''''        Set objExoticPowerPlant = Nothing
''''
''''    Case ClockWork, LeadAcidBattery, AdvancedBattery, Flywheel, RechargeablePowerCell, PowerCell
''''        Set objEnergyBank = New clsEnergyBank
''''        Components.Add objEnergyBank, CurrentItem
''''        Set objEnergyBank = Nothing
''''
''''    Case AntiMatterBay, CoalBunker, WoodBunker, StandardTank, lightTank, UltralightTank, StandardSelfSealingTank, lightSelfSealingTank, UltralightSelfSealingTank
''''        Set objFuelTank = New clsFuelTank
''''        Components.Add objFuelTank, CurrentItem
''''        Set objFuelTank = Nothing
''''
''''
''''    'Case Water, Wood, Coal, Gasoline, Diesel, AviationGas, JetFuel, Propane, LiquifiedNaturalGas, EthanolAlchohol, MethanolAlchohol, LiquidHydrogen, LiquidOxygen, Cadmium, MetalLOX, RocketFuel, AntiMatter
''''    '    Set objFuel = New clsFuel
''''    '    Components.Add objFuel, CurrentItem
''''    '    set objFuel = nothing
''''
''''    Case Snorkel
''''        Set objSnorkel = New clsSnorkel
''''        Components.Add objSnorkel, CurrentItem
''''        Set objSnorkel = Nothing
''''
''''    Case ElectricContactPower
''''        Set objElectricContactPower = New clsElectricContactPower
''''        Components.Add objElectricContactPower, CurrentItem
''''        Set objElectricContactPower = Nothing
''''
''''    Case LaserBeamedPowerReceiver, MaserBeamedPowerReceiver
''''        Set objBeamedPowerReceiver = New clsBeamedPowerReceiver
''''        Components.Add objBeamedPowerReceiver, CurrentItem
''''        Set objBeamedPowerReceiver = Nothing
''''
''''    Case NitrousOxideBooster
''''        Set objNitrousOxideBooster = New clsNitrousOxideBooster
''''        Components.Add objNitrousOxideBooster, CurrentItem
''''        Set objNitrousOxideBooster = Nothing
''''
''''    End Select
    
''''   TODO: So far its temporary deleted while i implement new clsFactory.CreateComponent.  If it works
''''   All of the below may be obsolte  MPJ Oct.2.2002
''''    'if we are loading the object from a File (or if copy/paste operation), then we need to exit here.
''''    If bFileLoadMode = True Then
''''        AddObject = True 'must return true since we didnt encounter any errors
''''        Exit Function
''''    End If
''''
''''    'add the initial required property values to the objects
''''    With Components.Item(CurrentItem)
''''        .Key = CurrentItem '<-- Todo: in each class, .key, .Parent, .Image, etc should be Friend "LET" and not Public "LET"
''''        .Parent = sParentKey   ' UNLESS making them "Friend" screws up the TLBINF32.DLL's ability to access these properties
''''        .Image = nImage
''''        .Datatype = Datatype
''''        .Description = sNodeText
''''        .CustomDescription = sNodeText
''''    End With
''''
''''    'TODO: Would be nice to actually perform this first... before adding object to treeview
''''    'TODO: In fact, this is even more desireable now that adding a node involves detecting if its even a "Component" node
''''    '      and not for instance, a profile or fuel link node.
''''    '      This wouldnt be quite as big a problem if we were passing OBJECT references since then we could use TypeOf.
''''
''''    'Query the object via its "LocationCheck" method to see if the user is
''''    'allowed to place the object in that locaton
''''    If Components.Item(CurrentItem).LocationCheck Then
''''        With Components.Item(CurrentItem)
''''            .Init
''''            .GetMatrixIndex
''''            .StatsUpdate
''''            .QueryParent
''''        End With
''''
''''        AddObject = True 'return the function results
''''        KeyManager.AddKeyChainKeys (CurrentItem) 'add Key References
''''    Else
''''        'we were unable to pass the LocationCheck for this component so we
''''        'must delete it from the collection and then tell the calling
''''        'prodcedure that the AddObjected failed so that it can handle any Treeview
''''        'clean up such as deleting a node
''''        Components.Remove CurrentItem
''''        AddObject = False 'return the function results
''''    End If
