Attribute VB_Name = "modProperties"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Const PROPERTY_HEADER = "HEADER"
'constants for the PropertyList datatypes
'todo: the CmpEdit will need to use these too... is this bas shared or should we rip out the constants and put them in a shareable module?
Private Const wdBool = 1
Private Const wdColor = 4
Private Const wdCurrency = 9
Private Const wdDate = 6
Private Const wdDefault = -1
Private Const wdDouble = 10
Private Const wdFile = 5
Private Const wdFont = 2
Private Const wdList = 3
Private Const wdNumber = 8
Private Const wdObject = 11
Private Const wdPicture = 7
Private Const wdText = 0
Private Const wdHeader = -99

Function formatCaption(ByVal lngDatatType As Long, ByVal lngUnitType As Long, s As String) As String
    s = " " & s
    formatCaption = s
End Function

Function formatValue(ByVal lngDatatype As Long, ByVal lngUnitType As Long, v As Variant) As Variant
    formatValue = v
End Function

Sub AddPCLproperty(ByRef vValue As Variant, ByRef oPropItem As cPropertyItem)
    Dim lngNewIndex As Long
    Dim lngDatatype As Long
    Dim lngUnitType As Long
    Dim sCaption As String
    Dim dblValue As Double
    Dim vItem As Variant
    Dim i As Long
    'Const SEPERATOR = "___________________"
    Const SEPERATOR = "===================="
    
    lngDatatype = oPropItem.Datatype ' proplist data type (e.g. list, float, number, text, etc)
    lngUnitType = oPropItem.UnitType ' gvd unit type used by the unit converter
    sCaption = oPropItem.Caption
    
    sCaption = formatCaption(lngDatatype, lngUnitType, sCaption)
    vValue = formatValue(lngDatatype, lngUnitType, vValue)
    'todo: when finally done, check out for places where i handle doubles and singles together to make sure its all correct.
    ' see notes below on some things to look out for
    
    ' NOTE: I use double's for most final stats, however I dont for table data since it would require double the amount of space
    ' and there are ALOT of tables.  Also used in our m_UserInput() variable which gets displayed directly here in the proplist.
    ' However, since the PLC1 doesnt support a "single" datatype, I must use the PLC1's double datatype
    ' which unfortunately will show approximation errors.  To fix this, we must convert the single to a string first and then assign that to the double.
    ' NOTE: Double to single doesnt cause these errors (as long as it doesnt overflow obviously) but single to double does.  Keep this in mind during
    ' stats calculations where we might be mixing the SINGLE's in our tables with the DOUBLE's required for the final answer.
    If lngDatatype = wdDouble Then
        dblValue = Val(CStr(vValue))
        frmDesigner.PLC1.AddItem dblValue, lngDatatype
    ElseIf lngDatatype = wdHeader Then
        With frmDesigner.PLC1
            .AddItem SEPERATOR, 0 ' change it back from -99 to 0 so that we can display the seperator line
            .ItemDisabledTextBold(.NewIndex) = True
        End With
    Else
        frmDesigner.PLC1.AddItem vValue, lngDatatype
    End If
        
    With frmDesigner.PLC1
        lngNewIndex = .NewIndex
        .ItemData(lngNewIndex) = lngDatatype
        .CaptionString(lngNewIndex) = sCaption
        .DescriptionString(lngNewIndex) = oPropItem.Notes
        .ItemDisabled(lngNewIndex) = oPropItem.ReadOnly
    End With
        
    'todo: how do we handle ammolists, guidancelists, or material/quality lists for armor?
    '      the armor layer would have to update its property item's list every time techlevel was changed.
    '      and if tech level was changed to somethign not supported by the current material, the material
    '     would default to first available at that techlevel. (e.g. cheap wood).  The only way to do this
    '     properly is use multiple matrices so that the values can be looked up based on the tech level and material.
    If lngDatatype = wdList Then
        vItem = oPropItem.List
        If IsArray(vItem) Then
            For i = LBound(vItem) To UBound(vItem)
                If vItem(i) <> "" Then
                    frmDesigner.PLC1.AddListItem (lngNewIndex), vItem(i)
                End If
            Next
        End If
    End If
    
    vValue = Empty
End Sub

Public Sub PropertyChanged(ByVal index As Long)
   Const LNG_LENGTH = 4
   Dim oProp As cPropertyItem
   Dim oNode As cINode
   Dim oDisplay As cIDisplay
   Dim lptr As Long
   Dim sClassname As String
   Dim lngInterfaceID As Long
   Dim lngRet As Long
   Debug.Print "Object Handle = " & frmDesigner.PLC1.Tag & " PLC1 INDEX = " & index
   
   On Error GoTo err
   lptr = Val(frmDesigner.PLC1.Tag)
   CopyMemory oNode, lptr, LNG_LENGTH

   If Not oNode Is Nothing Then
        sClassname = oNode.Classname
        Set oDisplay = oNode
                
        ' NOTE: its imperative that the order of properties in the proplist match the order of properties in the object
        Set oProp = oDisplay.getPropertyItemByIndex(index)
        lngInterfaceID = oProp.interfaceid
        
        If Not oProp.ReadOnly Then
            ' todo: i think this select case is "ok"... anyway to modify it or perhaps consolidate the code
            '       with modProperties:PropertiesShow() ??  Actually probably not, it uses "vbGet" and this uses
            '       vbLet.  Actually what we should do then is move this out into a seperate function
            Debug.Print "modProperties:PropertyChanged() -- InterfaceID = " & lngInterfaceID & " PropertyName = " & oProp.CallByName
             Select Case lngInterfaceID
                Case INTERFACE_COMPONENT
                    Dim oComponent As cIComponent
                    Set oComponent = oNode
                    CallByName oComponent, oProp.CallByName, VbLet, frmDesigner.PLC1.value(index)

                Case INTERFACE_CONTAINER
                    Dim oContainer As cIContainer
                    Set oContainer = oNode
                    CallByName oContainer, oProp.CallByName, VbLet, frmDesigner.PLC1.value(index)
                Case INTERFACE_VEHICLE_DESCRIPTION
                    Dim oVehicle As cVehicle
                    Set oVehicle = oNode
                    CallByName oVehicle.Description, oProp.CallByName, VbLet, frmDesigner.PLC1.value(index)
                Case INTERFACE_VEHICLE_VERSION
                    CallByName oVehicle.version, oProp.CallByName, VbLet, frmDesigner.PLC1.value(index)
                    
                Case INTERFACE_VEHICLE_AUTHOR
                    CallByName oVehicle.author, oProp.CallByName, VbLet, frmDesigner.PLC1.value(index)
                    
                Case INTERFACE_NODE
                    CallByName oNode, oProp.CallByName, VbLet, frmDesigner.PLC1.value(index)
                Case INTERFACE_DISPLAY
                    CallByName oDisplay, oProp.CallByName, VbLet, frmDesigner.PLC1.value(index)
                Case INTERFACE_BUILD
                    Dim oBuild As cIBuild
                    Set oBuild = oNode
                    If oProp.CallBytype = 1 Then
                        ' since this is a function call, we have to modify the name to append "get"
                        ' also, if its a wdList, we need to pass the subscript of the selected list item
                        If oProp.Datatype = wdList Then
                            lngRet = oProp.getSelectionIndexFromValue(frmDesigner.PLC1.value(index))
                            If lngRet <> -1 Then
                                CallByName oBuild, "set" & oProp.CallByName, VbMethod, oProp.Subscript, lngRet
                            Else
                                ' an error and there shouldnt be any
                                MsgBox "modProperties:PropertyChanged() -- Invalid wdList subscript"
                            End If
                        ' else its userinput and we pass the actual value
                        ElseIf oProp.Datatype = wdDouble Then
                            CallByName oBuild, "set" & oProp.CallByName, VbMethod, oProp.Subscript, frmDesigner.PLC1.value(index)
                        Else ' we should never reach here
                            
                        End If
                    Else
                        CallByName oBuild, oProp.CallByName, VbLet, frmDesigner.PLC1.value(index)
                    End If
                Case INTERFACE_SURFACE
                    Dim oSurface As cSurface
                    Set oSurface = oNode
                    lngRet = oProp.getSelectionIndexFromValue(frmDesigner.PLC1.value(index))
                    If lngRet <> -1 Then
                        CallByName oSurface, oProp.CallByName, VbLet, lngRet
                    Else
                        ' an error and there shouldnt be any
                        MsgBox "modProperties:PropertyChanged() -- Invalid wdList subscript"
                    End If
                            
                Case Else
                    Debug.Print "modProperties:PropertyChanged() -- Class Interface Not Supported."
            End Select
        Else
             InfoPrint 1, "The property '" & oProp.Caption & "' is Read Only."
        End If
        
        Set oProp = Nothing
   End If
   CopyMemory oNode, 0&, LNG_LENGTH
   Set oDisplay = Nothing
   DoEvents
'    UpdateVehicle todo: all of these may be obosolete under new code base EXCEPT when user hits F5 or when forcing recalc after laoding saved vehicle
     
    '//place the cell back into its original spot. Todo: why is this needed?
    frmDesigner.PLC1.ListIndex = index
    p_bChangedFlag = True ' JAW 2000.05.07 change has been made in vehicle/component
   Exit Sub
err:
    Debug.Print "modProperties:PropertyChanged() -- Error #" & err.Number & " " & err.Description
    If Not oNode Is Nothing Then
        CopyMemory oNode, 0&, LNG_LENGTH
        Set oDisplay = Nothing
        Set oProp = Nothing
   End If
   
'    Todo: still need to properly handle these types.... perhaps new datatype like wdNote (signifies text greater than X chars?)
'    If frmDesigner.PLC1.DescriptionString(Index) = "StationFunction" Then
'        Load frmNotes
'        frmNotes.Tag = "crewstation"
'        frmNotes.Show vbModal, frmDesigner
'        Set frmNotes = Nothing
'

'
End Sub
Public Sub Properties_Show(ByVal hNode As Long)
    If hNode <= 0 Then Exit Sub
       
    Dim oNode As Vehicles.cINode
    Dim oDisplay As cIDisplay
    Dim oPropItem As Vehicles.cPropertyItem
    Dim vValue As Variant
    Dim lngInterfaceID As Long
    Const LNG_LENGTH = 4
    Dim index As Long
    On Error GoTo errDefault
    
    With frmDesigner.PLC1
        .Clear
        .ShowDescription = True
        .Tag = hNode  'CRITICAL - needed so that when proplist attributes for a component are changed, the PLC1 code knows which item is being referenced
    End With
    
    CopyMemory oNode, hNode, LNG_LENGTH '<--- every component in the _Tree_ MUST implement cINode because thats what this pointer is for (if its not going to be rendered in the tree, it doesnt need that interface
    Set oDisplay = oNode               '<-- every component must also obviously implement cIDisplay
    
    If Not oDisplay Is Nothing Then
    Set oPropItem = oDisplay.getfirstpropertyitem
        Do While Not oPropItem Is Nothing
            If Not oPropItem.Caption = PROPERTY_HEADER Then
                On Error GoTo errVarName 'todo: this helps us get past bugs while under development, where properties dont exist so callbyname fails
                'NOTE: since the index location of the property in the PLC1  MUST correspond to the  index of the property in the array in our oNode
                ' we should fill in a blank line for any property that fails to properly load.

                lngInterfaceID = oPropItem.interfaceid
                Debug.Print "modProperties:Properties_Show() -- InterfaceID = " & lngInterfaceID & " PropertyName = " & oPropItem.CallByName
                Select Case lngInterfaceID
                'todo: this entire select case needs to be moved to  a seperate function
                'todo: and wouldnt it just be better to call oNode.ClassName and then do the select case by TypeName??
                '      actually, I dont think I can since with composite objects (like armor inside a component) there will exist
                '      multiple interface ID's.  Using typename will not allow us to switch between interfaces.
                '      The real question is, is there a way to get rid of having a huge select case statement?
                    Case INTERFACE_NODE
                        vValue = CallByName(oNode, oPropItem.CallByName, VbGet)
                    
                    Case INTERFACE_COMPONENT
                        ' NOTE: keep this twoard top of select case since its an often used case
                        Dim oComponent As cIComponent
                        Set oComponent = oNode
                        vValue = CallByName(oComponent, oPropItem.CallByName, VbGet)

                    Case INTERFACE_ARMOR
                        Dim oArmor As cArmor
                        Set oArmor = oNode
                        vValue = CallByName(oArmor, oPropItem.CallByName, VbGet)
                        
                    Case INTERFACE_CONTAINER
                        Dim oContainer As cIContainer
                        Set oContainer = oNode
                        vValue = CallByName(oContainer, oPropItem.CallByName, VbGet)
                    Case INTERFACE_VEHICLE_DESCRIPTION
                        Dim oVehicle As cVehicle
                        Set oVehicle = oNode
                        vValue = CallByName(oVehicle.Description, oPropItem.CallByName, VbGet)
                                                
                    Case INTERFACE_VEHICLE_VERSION
                       ' Dim oVehicle As cVehicle
                        Set oVehicle = oNode
                        vValue = CallByName(oVehicle.version, oPropItem.CallByName, VbGet)
                        
                    Case INTERFACE_VEHICLE_AUTHOR
                       ' Dim oVehicle As cVehicle
                        Set oVehicle = oNode
                        vValue = CallByName(oVehicle.author, oPropItem.CallByName, VbGet)
                    
                    Case INTERFACE_BUILD
                        Dim oBuild As cIBuild
                        Set oBuild = oNode
                        If oPropItem.CallBytype = 1 Then
                            vValue = CallByName(oBuild, "get" & oPropItem.CallByName, VbMethod, oPropItem.Subscript)
                            ' if its user input, we display the value
                            If oPropItem.Datatype = wdDouble Then
                                'vValue = vValue
                            ' else its an option and we use the returned index value to find the string represenation for the selection
                            ElseIf oPropItem.Datatype = wdList Then
                                vValue = oPropItem.ListItem(vValue)
                            Else
                                MsgBox "modProperties.Properties_Show() -- Error: Undefined property type."
                            End If
                        Else
                            vValue = CallByName(oBuild, oPropItem.CallByName, VbGet)
                        End If
                            
                    Case INTERFACE_SURFACE
                        Dim oSurface As cSurface
                        Set oSurface = oNode
                    
                        'todo: im potentially going to wind up with the same style If/else block for every interface that has a wdList
                        '      is there another way to design this?  Well, maybe its not too many interfaces afterall?  We will see
                        If oPropItem.Datatype = wdList Then
                            vValue = CallByName(oSurface, oPropItem.CallByName, VbGet)
                            vValue = oPropItem.ListItem(vValue)
                        Else
                            vValue = CallByName(oSurface, oPropItem.CallByName, VbGet)
                        End If
                    
                    Case Else
                        'problem
                        InfoPrint 1, "modProperties:Properties_Show() -- Unsupported Class Interface ID '" & lngInterfaceID & "'  Cannot list property '" & oPropItem.Caption & "'"
                End Select
            End If
            
            On Error GoTo errDefault
            'NOTE: this property must get added here regardless of whether there was a problem accessing its value
            AddPCLproperty vValue, oPropItem
            ' get the next one
            Set oPropItem = oDisplay.getnextpropertyitem
        Loop
        CopyMemory oNode, 0&, LNG_LENGTH
        Set oPropItem = Nothing
        Set oDisplay = Nothing
    End If
    
    'set the column width of the proplist to always be half the total width
    'todo: this should be done on event when the width of this control changes
   frmDesigner.PLC1.ColumnWidth = frmDesigner.PLC1.Width / (2 * Screen.TwipsPerPixelX)
    Exit Sub

errVarName:
    Debug.Print "modProperties.Properties_Show() -- Could not get value for Variable Name '" & oPropItem.CallByName & "'"
    Resume Next
errDefault:
    Debug.Print "modProperties:Properties_Show -- Error #" & err.Number & " " & err.Description
    If Not oNode Is Nothing Then
        CopyMemory oNode, 0&, LNG_LENGTH
    End If
    Set oPropItem = Nothing
    Set oDisplay = Nothing
End Sub

Public Sub DisplayPrintOutput()
    
    Dim sKey As String
    Dim lngDatatype As Long
    Dim sParentKey As String
    'todo: This could should actually be integrated with the Show_Properties since it has the same
    '      task of determine what type of node we're dealing with
    sKey = p_ActiveNode.Key
    lngDatatype = p_ActiveNode.Datatype
    sParentKey = p_ActiveNode.Parent
    'display the print output in the status bar for this node
    Select Case lngDatatype
        Case Body
            'frmDesigner.StatusBar1.Panels(1).text = m_oCurrentVeh.Body.PrintOutput
        Case VEHICLE_NODE
            ' do nothing
        ' is a component
        Case PERFORMANCE_NODE
        Case CREW_NODE
        Case POWERSYSTEMS_NODE
            If sParentKey = POWERSYSTEMS_KEY Then
                'frmDesigner.StatusBar1.Panels(1).Text = m_oCurrentVeh.Profiles(sKey).Description
            End If
        Case WEAPON_LINKS_NODE
            If sParentKey = WEAPON_LINKS_KEY Then
                'frmDesigner.StatusBar1.Panels(1).Text = m_oCurrentVeh.WeaponProfiles(sKey).Description & "  --  " & Format(m_oCurrentVeh.WeaponProfiles(sKey).Cost, vbCurrency)
            End If
        Case FUELSYSTEMS_NODE
            If sParentKey = FUELSYSTEMS_KEY Then
                'frmDesigner.StatusBar1.Panels(1).Text = m_oCurrentVeh.Profiles(sKey).Description
            End If
        Case PERFORMANCEAIR To PERFORMANCESPACE
            'frmDesigner.StatusBar1.Panels(1).Text = m_oCurrentVeh.PerformanceProfiles(sKey).Description
            
        Case Else
            ' regular components
            'todo: uncomment
           ' frmDesigner.StatusBar1.Panels(1).text = m_oCurrentVeh.Components(sKey).PrintOutput
    End Select
End Sub

Private Sub ShowPropsForPerformanceProfile()

    PopulateCheckList m_oCurrentVeh.ActiveCheckListType
    Dim sKey As String
    sKey = m_oCurrentVeh.ActiveCheckList
    If sKey <> "" Then
        If m_oCurrentVeh.ActiveCheckListType = WEAPON_CHECKLIST Then
            ShowPropsForWeaponLink sKey
        Else
            ShowPropsForPerformance sKey
        End If
    End If
            
End Sub
Private Sub ShowPropsForPowerProfile()
    Dim sKey As String
    sKey = m_oCurrentVeh.ActiveProfile
    If sKey <> "" Then
        m_oCurrentVeh.Profiles(sKey).Show
    
        If m_oCurrentVeh.ActiveProfiletype = FUEL_PROFILE Then
            Call ShowLinks(FUEL_PROFILE)
        Else
            Call ShowLinks(POWER_PROFILE)
        End If
    End If
End Sub
Private Sub ShowPropsForDescription()
    'these settings are available off of vehicle node
End Sub

Private Sub ShowPropsForOptions()
    'dont need this... this shows up for vehicle node
End Sub

Private Sub ShowPropsForCrew()

    With m_oCurrentVeh.crew
      AddPCLproperty "Settings", "", wdText, PROPERTY_HEADER
      AddPCLproperty "Use Recommended Crew", .UseRecommendedCrew, wdBool, "UseRecommendedCrew"
      AddPCLproperty "Occupancy", .Occupancy, wdList, "Occupancy", "short", "long"
      AddPCLproperty "Number of Shifts", .numshifts, wdList, "NumShifts", 1, 2, 3, 4, 5, 6, 7, 8, 9
      AddPCLproperty "Military Vehicle", .MilitaryVehicle, wdBool, "MilitaryVehicle"
      
      AddPCLproperty "Crew Quantities", "", wdText, PROPERTY_HEADER
      AddPCLproperty "Captains", .numcaptains, wdNumber, "NumCaptains"
      AddPCLproperty "Officers", .NumOfficers, wdNumber, "NumOfficers"
      AddPCLproperty "Crew Station Operators", .NumCrewStationOperators, wdNumber, "NumCrewStationOperators"
      AddPCLproperty "Weapon Loaders", .NumWeaponLoaders, wdNumber, "NumWeaponLoaders"
      
      AddPCLproperty "Rowers", .NumRowers, wdNumber, "NumRowers"
      AddPCLproperty "Sailors", .NumSailors, wdNumber, "NumSailors"
      AddPCLproperty "Riggers", .NumRiggers, wdNumber, "NumRiggers"
      AddPCLproperty "Fuel Stokers", .NumFuelStokers, wdNumber, "NumFuelStokers"
      AddPCLproperty "Mechanics", .NumMechanics, wdNumber, "NumMechanics"
      AddPCLproperty "Service Crewmen", .NumServiceCrewmen, wdNumber, "NumServiceCrewmen"
      AddPCLproperty "Medics", .NumMedics, wdNumber, "NumMedics"
      AddPCLproperty "Scientists", .NumScientists, wdNumber, "NumScientists"
      AddPCLproperty "Auxiliary Vehicle Crew", .NumAuxiliaryVehicleCrew, wdNumber, "NumAuxiliaryVehicleCrew"
      AddPCLproperty "Stewards", .NumStewards, wdNumber, "NumStewards"
      
      AddPCLproperty "Passenger Quantities", "", wdText, PROPERTY_HEADER
      AddPCLproperty "Luxury", .NumLuxury, wdNumber, "NumLuxury"
      AddPCLproperty "First Class", .NumFirstClass, wdNumber, "NumFirstClass"
      AddPCLproperty "Second Class", .NumSecondClass, wdNumber, "NumSecondClass"
      AddPCLproperty "Steerage", .NumSteerage, wdNumber, "NumSteerage"
      
      AddPCLproperty "Stats", "", wdText, PROPERTY_HEADER
      AddPCLproperty "Total Crew + Passengers", .TotalNumberCrewPassengers, wdNumber, "TotalNumberCrewPassengers"
    End With
End Sub
 
Private Sub ShowPropsForSurface()
    
End Sub

Private Sub ShowPropsForStats()
   
With m_oCurrentVeh.Description
    Dim vCategories() As Variant
    Dim vSubCategories() As Variant
    Call LoadCategories(vCategories)
    Call LoadSubCategories("Wheeled", vSubCategories)
     
    On Error Resume Next '<-- todo: this is because when you first create a vehicle, the two vCategories() and vSubCategories lines that follow will raise an error
    'Vehicle Description and Authoring Information
    AddPCLproperty "Vehicle Description", "", wdText, PROPERTY_HEADER
    AddPCLproperty "Name", .NickName, wdText, "NickName"
    AddPCLproperty "Class", .Classname, wdText, "ClassName"
    'todo: Hrm... categories suck.  They were implemented originally with the intent that they could be used
    ' to filter submissions to the website.  However, i think a better way to handle this is for the user to
    ' select the cat/subcat WHEN they want to upload it.  Primarily because these categories can change
    ' on the website, and users wont always have proper categories in GVD
    AddPCLproperty "Category", .Category, wdList, "Category", vCategories()  'todo: this one (see above)
    AddPCLproperty "Sub Category", .subcategory, wdList, "subcategory", vSubCategories() 'todo: and this one (see above)
      
    AddPCLproperty "Description", .VehicleDescription, wdText, "VehicleDescription"
    AddPCLproperty "Details", .Details, wdText, "Details"
    AddPCLproperty "Vision", .Vision, wdText, "Vision"
    AddPCLproperty "Header", .Header, wdText, "Header"
    AddPCLproperty "Footer", .Footer, wdText, "Footer"
    AddPCLproperty "VehicleImageFileName", .VehicleImageFileName, wdText, "VehicleImageFileName"
       
    AddPCLproperty "Versioning", "", wdText, PROPERTY_HEADER
    AddPCLproperty "Auto Increment Version", .blnAutoIncrement, wdBool, "blnAutoIncrement"
    AddPCLproperty "Vehicle Version", .version, wdText, "version" 'todo: make sure its read only
      
      
    AddPCLproperty "Author Info", "", wdText, PROPERTY_HEADER
    AddPCLproperty "Name", .author, wdText, "Author"
    AddPCLproperty "Email", .email, wdText, "Email"
    AddPCLproperty "Website", .url, wdText, "Url"
End With

 With m_oCurrentVeh.Options
      
      AddPCLproperty "Miscellaneous", "", wdText, PROPERTY_HEADER
      AddPCLproperty "Vehicle Crafstmanship", .Quality, wdList, "Quality", "standard", "cheap", "fine", "very fine"
      AddPCLproperty "RollStabilizers", .RollStabilizers, wdBool, "RollStabilizers"
      AddPCLproperty "Convertible", .Convertible, wdList, "Convertible", "none", "hardtop", "ragtop"
      AddPCLproperty "UseHardpointMountedWeights", .UseHardpointMountedWeights, wdBool, "UseHardpointMountedWeights"
           
      AddPCLproperty "Payload Settings", "", wdText, PROPERTY_HEADER
      AddPCLproperty "Use Default Weights", .RecommendedPayload, wdBool, "RecommendedPayload"
      AddPCLproperty "Per Person Weight", .PerPersonWeight, wdNumber, "PerPersonWeight"
      AddPCLproperty "Per Cargo Weight", .PerCargoWeight, wdNumber, "PerCargoWeight"
      
      AddPCLproperty "Access Space", "", wdText, PROPERTY_HEADER
      AddPCLproperty "Use Recommended Modifier?", .RecommendedAccessSpace, wdBool, "RecommendedAccessSpace"
      AddPCLproperty "Volume Modifier", .AccessSpaceVolumeMod, wdList, "AccessSpaceVolumeMod", 0, 0.25, 0.5, 0.75, 1, 1.25, 1.5, 1.75, 2
        
      AddPCLproperty "Attachments", "", wdText, PROPERTY_HEADER
      AddPCLproperty "Pin", .Pin, wdList, "Pin", "none", "standard", "Explosive"
      AddPCLproperty "Ram", .Ram, wdBool, "ram"
      AddPCLproperty "Bulldozer", .Bulldozer, wdBool, "bulldozer"
      AddPCLproperty "Plow", .Plow, wdBool, "Plow"
      AddPCLproperty "Hitch", .Hitch, wdBool, "hitch"
    End With
    
 With m_oCurrentVeh.surface
       AddPCLproperty "Hull and Hydro Options", "", wdText, PROPERTY_HEADER
       AddPCLproperty "Streamlining", .StreamLining, wdList, "Streamlining", "none", "fair", "good", "very good", "superior", "excellent", "radical"
       AddPCLproperty "Floatation Hull", .FloatationHull, wdBool, "floatationhull"
       AddPCLproperty "Submerisible Hull (TL5)", .Submersible, wdBool, "Submersible"
       AddPCLproperty "Hydrodynamic Lines", .HydrodynamicLines, wdList, "Hydrodynamiclines", "none", "mediocre", "average", "fine", "very fine", "submarine"
       'AddPCLproperty "Roll Stabilizers (TL7)", .RollStabilizers, wdBool, "rollstabilizers"
       AddPCLproperty "Waterproof", .WaterProof, wdBool, "waterproof"
       AddPCLproperty "Sealed (TL5)", .Sealed, wdBool, "Sealed"
       AddPCLproperty "Cata/Tri(maran)", .CataTrimaran, wdList, "catatrimaran", "none", "catamaran", "trimaran"

       
       
       AddPCLproperty "Concealment", "", wdText, PROPERTY_HEADER
       AddPCLproperty "Camouflage", .Camouflage, wdBool, "Camouflage"
       AddPCLproperty "Infrared Cloaking (TL7)", .infraredcloaking, wdList, "InfraredCloaking", "none", "basic", "radical"
       AddPCLproperty "Emission Cloaking (TL8)", .EmissionCloaking, wdList, "EmissionCloaking", "none", "basic", "radical"
       AddPCLproperty "Sound Baffling (TL7)", .SoundBaffling, wdList, "SoundBaffling", "none", "basic", "radical"
       AddPCLproperty "Stealth (TL7)", .stealth, wdList, "Stealth", "none", "basic", "radical"
       AddPCLproperty "Liquid Crystal Skin (TL8)", .LiquidCrystal, wdBool, "LiquidCrystal"
       AddPCLproperty "PsiShielding (TL8)", .PsiShielding, wdBool, "PsiShielding"
       AddPCLproperty "Chameleon", .Chameleon, wdList, "Chameleon", "none", "basic", "instant", "intruder"
    
       AddPCLproperty "Magic Levitation", "", wdText, PROPERTY_HEADER
       AddPCLproperty "Enabled", .bMagicLevitation, wdBool, "bMagicLevitation"
       AddPCLproperty "Energy Cost Per Pound", .MagicLevitationEnergyCostPerPound, wdDouble, "MagicLevitationEnergyCostPerPound"
       
       AddPCLproperty "Antigravity Coating", "", wdText, PROPERTY_HEADER
       AddPCLproperty "Enabled", .bAntigravityCoating, wdBool, "bAntigravityCoating"
       AddPCLproperty "Cost Per Sq ft", .AntigravityCoatingCostPerSquareFoot, wdDouble, "AntigravityCoatingCostPerSquareFoot"
       AddPCLproperty "Surface Area Useage", .AntigravityCoatingSurfaceAreaUseage, wdList, "AntigravityCoatingSurfaceAreaUseage", "Body", "Vehicle"
      
       AddPCLproperty "Super Science Coating", "", wdText, PROPERTY_HEADER
       AddPCLproperty "Enabled", .bSuperScienceCoating, wdBool, "bSuperScienceCoating"
       AddPCLproperty "Cost Per Sq ft", .SuperScienceCoatingCostPerSquareFoot, wdDouble, "SuperScienceCoatingCostPerSquareFoot"
       AddPCLproperty "Surface Area Useage", .SuperScienceCoatingSurfaceAreaUseage, wdList, "SuperScienceCoatingSurfaceAreaUseage", "Body", "Vehicle"
     
    End With
    
'03/24/02 - We dont need to display stats here... it just makes for too much scrolling
'With m_oCurrentVeh.Stats
'    'update displays for Price, Health and SizeModifier ,Volume, Weight and Emptyweigh
'    AddPCLproperty "Statistics", "", wdText, "Disabled"
'    AddPCLproperty "Price", "$" & Format(.TotalPrice, "standard"), wdText, "Disabled"
'
'
'    AddPCLproperty "Health", .StructuralHealth & " HT", wdText, "Disabled"
'    AddPCLproperty "SizeMod", .SizeModifier, wdText, "Disabled"
'    AddPCLproperty "Volume", Format(.TotalVolume, "standard") & " cu ft", wdText, "Disabled"
'    AddPCLproperty "Area", Format(.totalsurfacearea, "standard") & " sq ft", wdText, "Disabled"
'    AddPCLproperty "Empty Wt", Format(.EmptyWeight, "standard") & " lbs", wdText, "Disabled"
'    AddPCLproperty "Empty Mass", Format(.EmptyWeight / 2000, "standard") & " tons", wdText, "Disabled"
'    AddPCLproperty "Loaded Wt", Format(.LoadedWeight, "standard") & " lbs", wdText, "Disabled"
'    AddPCLproperty "Loaded Mass", Format(.LoadedMass, "standard") & " tons", wdText, "Disabled"
'    AddPCLproperty "+Hardpoint Wt", Format(.HLoadedWeight, "standard") & " lbs", wdText, "Disabled"
'    AddPCLproperty "+Hardpoint Mass", Format(.HLoadedMass, "standard") & " tons", wdText, "Disabled"
'    AddPCLproperty "Submerged Wt", Format(.SubmergedWeight, "standard") & " lbs", wdText, "Disabled"
'    AddPCLproperty "Submerged Mass", Format(.SubmergedMass, "standard") & " tons", wdText, "Disabled"
'    AddPCLproperty "Power Output", Format(.TotalGeneratedPower, "standard") & " kW", wdText, "Disabled"
'    AddPCLproperty "Power Consumption", Format(.TotalContinuousPowerConsumption, "standard") & " kW", wdText, "Disabled"
'    AddPCLproperty "Flotation Rating", .FloatationRating, wdText, "Disabled"
'End With
'



 'todo: StatusBar1.Panels(2).Text = "Added Wt: " & Format(.OptionsWeight, "standard") & " lbs"
'    StatusBar1.Panels(1).Text = "Added Cost: $" & Format(.OptionsCost, "standard")
'    StatusBar1.Panels(3).Text = "Internal Payload Wt: " & Format(.UsualInternalPayload, "standard") & " lbs"
    

End Sub

Public Sub ShowPropsForPerformance(ByVal Key As String)
    
    'PerformanceType is stored in  m_oCurrentVeh.PerformanceProfiles(Key).Datatype
    
    With m_oCurrentVeh.PerformanceProfiles(Key)

        Select Case .Datatype
        'JAW 2000.06.18
        'Added takeoff/land
            
            Case PERFORMANCELEG
                AddPCLproperty "Legged Performance", "", wdText, PROPERTY_HEADER
                
                AddPCLproperty "Thrust Options", "", wdText, PROPERTY_HEADER
                AddPCLproperty "PercentThrust", .percentthrust, wdNumber, "PercentThrust"
                AddPCLproperty "TreatTiltRotorsAsPropellers", .TreatTiltRotorsAsPropellers, wdBool, "TreatTiltRotorsAsPropellers"
                AddPCLproperty "AfterBurnersOn", .AfterBurnersOn, wdBool, "AfterBurnersOn"
                
                
                AddPCLproperty "Streamlining", "", wdText, PROPERTY_HEADER
                AddPCLproperty "HardPointsOn", .HardPointsOn, wdBool, "HardPointsOn"
                AddPCLproperty "WheelsSkidsExtended", .WheelsSkidsExtended, wdBool, "WheelsSkidsExtended"
                AddPCLproperty "PopTurretsExtended", .PopTurretsExtended, wdBool, "PopTurretsExtended"
                
                AddPCLproperty "Weight Percentages", "", wdText, PROPERTY_HEADER
                AddPCLproperty "% Crew", .PercentCrewWeight, wdNumber, "PercentCrewWeight"
                AddPCLproperty "% Fuel", .PercentFuelWeight, wdNumber, "PercentFuelWeight"
                AddPCLproperty "% Cargo", .PercentCargoWeight, wdNumber, "PercentCargoWeight"
                AddPCLproperty "% Hardpoints/Bays Load", .PercentHardpointWeight, wdNumber, "PercentHardpointWeight"
                AddPCLproperty "% Provisions", .PercentProvisionWeight, wdNumber, "PercentProvisionWeight"
                AddPCLproperty "% Ammunitions", .PercentAmmunitionWeight, wdNumber, "PercentAmmunitionWeight"
                AddPCLproperty "% PercentAuxVehicleWeight", .PercentAuxVehicleWeight, wdNumber, "PercentAuxVehicleWeight"
                
                AddPCLproperty "Statistics", "", wdText, PROPERTY_HEADER
                AddPCLproperty "Total Drivetrain Power", Format(.gtotalmotivepower, "standard") & " kW", wdText, "Disabled"
                AddPCLproperty "gSpeed", Format(.gTopSpeed, "standard") & " mph", wdText, "Disabled"
                AddPCLproperty "gOffRd", Format(.gOffRoad, "standard") & " mph", wdText, "Disabled"
                AddPCLproperty "gAccel", .gAcceleration & " mph/s", wdText, "Disabled"
                AddPCLproperty "gDecel", .gDeceleration & " mph/s", wdText, "Disabled"
                AddPCLproperty "gSR", .gStability, wdText, "Disabled"
                AddPCLproperty "gMR", .gManeuverability, wdText, "Disabled"
                AddPCLproperty "gP", Format(.gPressure, "standard"), wdText, "Disabled"
                AddPCLproperty "gPDescr", .gPressureDescription, wdText, "Disabled"
                AddPCLproperty "Advisory", .Advisory, wdText, "Disabled"
                
            Case PERFORMANCETRACK
                AddPCLproperty "Tracked Performance", "", wdText, PROPERTY_HEADER
                AddPCLproperty "gSpeed", Format(.gTopSpeed, "standard") & " mph", wdText, "Disabled"
                AddPCLproperty "gOffRd", Format(.gOffRoad, "standard") & " mph", wdText, "Disabled"
                AddPCLproperty "gAccel", .gAcceleration & " mph/s", wdText, "Disabled"
                AddPCLproperty "gDecel", .gDeceleration & " mph/s", wdText, "Disabled"
                AddPCLproperty "gSR", .gStability, wdText, "Disabled"
                AddPCLproperty "gMR", .gManeuverability, wdText, "Disabled"
                AddPCLproperty "gP", Format(.gPressure, "standard"), wdText, "Disabled"
                AddPCLproperty "gPDescr", .gPressureDescription, wdText, "Disabled"
                AddPCLproperty "Advisory", .Advisory, wdText, "Disabled"
                
            Case PERFORMANCEWHEEL
                AddPCLproperty "Wheeled Performance", "", wdText, PROPERTY_HEADER
                AddPCLproperty "gSpeed", Format(.gTopSpeed, "standard") & " mph", wdText, "Disabled"
                AddPCLproperty "gOffRd", Format(.gOffRoad, "standard") & " mph", wdText, "Disabled"
                AddPCLproperty "gAccel", .gAcceleration & " mph/s", wdText, "Disabled"
                AddPCLproperty "gDecel", .gDeceleration & " mph/s", wdText, "Disabled"
                AddPCLproperty "gSR", .gStability, wdText, "Disabled"
                AddPCLproperty "gMR", .gManeuverability, wdText, "Disabled"
                AddPCLproperty "gP", Format(.gPressure, "standard"), wdText, "Disabled"
                AddPCLproperty "gPDescr", .gPressureDescription, wdText, "Disabled"
                AddPCLproperty "Advisory", .Advisory, wdText, "Disabled"
                
            Case PERFORMANCEFLEX
                AddPCLproperty "Flexibody Performance", "", wdText, PROPERTY_HEADER
                AddPCLproperty "gSpeed", Format(.gTopSpeed, "standard") & " mph", wdText, "Disabled"
                AddPCLproperty "gOffRd", Format(.gOffRoad, "standard") & " mph", wdText, "Disabled"
                AddPCLproperty "gAccel", .gAcceleration & " mph/s", wdText, "Disabled"
                AddPCLproperty "gDecel", .gDeceleration & " mph/s", wdText, "Disabled"
                AddPCLproperty "gSR", .gStability, wdText, "Disabled"
                AddPCLproperty "gMR", .gManeuverability, wdText, "Disabled"
                AddPCLproperty "gP", Format(.gPressure, "standard"), wdText, "Disabled"
                AddPCLproperty "gPDescr", .gPressureDescription, wdText, "Disabled"
                AddPCLproperty "Advisory", .Advisory, wdText, "Disabled"
                
            Case PERFORMANCESKID
                AddPCLproperty "Skid Performance", "", wdText, PROPERTY_HEADER
                AddPCLproperty "gSpeed", Format(.gTopSpeed, "standard") & " mph", wdText, "Disabled"
                AddPCLproperty "gOffRd", Format(.gOffRoad, "standard") & " mph", wdText, "Disabled"
                AddPCLproperty "gAccel", .gAcceleration & " mph/s", wdText, "Disabled"
                AddPCLproperty "gDecel", .gDeceleration & " mph/s", wdText, "Disabled"
                AddPCLproperty "gSR", .gStability, wdText, "Disabled"
                AddPCLproperty "gMR", .gManeuverability, wdText, "Disabled"
                AddPCLproperty "gP", Format(.gPressure, "standard"), wdText, "Disabled"
                AddPCLproperty "gPDescr", .gPressureDescription, wdText, "Disabled"
                AddPCLproperty "Advisory", .Advisory, wdText, "Disabled"
                
            Case PERFORMANCEAIR
                AddPCLproperty "Air Performance", "", wdText, PROPERTY_HEADER
                AddPCLproperty "Thrust Options", "", wdText, PROPERTY_HEADER
                AddPCLproperty "PercentThrust", .percentthrust, wdNumber, "PercentThrust"
                AddPCLproperty "TreatTiltRotorsAsPropellers", .TreatTiltRotorsAsPropellers, wdBool, "TreatTiltRotorsAsPropellers"
                AddPCLproperty "AfterBurnersOn", .AfterBurnersOn, wdBool, "AfterBurnersOn"
                
                
                AddPCLproperty "Streamlining", "", wdText, PROPERTY_HEADER
                AddPCLproperty "HardPointsOn", .HardPointsOn, wdBool, "HardPointsOn"
                AddPCLproperty "WheelsSkidsExtended", .WheelsSkidsExtended, wdBool, "WheelsSkidsExtended"
                AddPCLproperty "PopTurretsExtended", .PopTurretsExtended, wdBool, "PopTurretsExtended"
                
                AddPCLproperty "Weight Percentages", "", wdText, PROPERTY_HEADER
                AddPCLproperty "% Crew", .PercentCrewWeight, wdNumber, "PercentCrewWeight"
                AddPCLproperty "% Fuel", .PercentFuelWeight, wdNumber, "PercentFuelWeight"
                AddPCLproperty "% Cargo", .PercentCargoWeight, wdNumber, "PercentCargoWeight"
                AddPCLproperty "% Hardpoints/Bays Load", .PercentHardpointWeight, wdNumber, "PercentHardpointWeight"
                AddPCLproperty "% Provisions", .PercentProvisionWeight, wdNumber, "PercentProvisionWeight"
                AddPCLproperty "% Ammunitions", .PercentAmmunitionWeight, wdNumber, "PercentAmmunitionWeight"
                AddPCLproperty "% PercentAuxVehicleWeight", .PercentAuxVehicleWeight, wdNumber, "PercentAuxVehicleWeight"
                
                AddPCLproperty "Statistics", "", wdText, PROPERTY_HEADER
                
                'frmDesigner.lblperformance(0).Caption =  "Can Fly?" & vbTab & .aCanFly
                AddPCLproperty "Thrust", Format(.aMotiveThrust, "standard") & " lbs", wdText, "Disabled"
                AddPCLproperty "Static Lift", Format(.staticlift, "standard") & " lbs", wdText, "Disabled"
                AddPCLproperty "Drag", .aDrag, wdText, "Disabled"
                AddPCLproperty "Speed", Format(.aTopSpeed, "standard") & " mph", wdText, "Disabled"
                AddPCLproperty "Stall Speed", Format(.aStallSpeed, "standard") & " mph", wdText, "Disabled"
                AddPCLproperty "aAccel", Format(.aAcceleration, "standard") & " mph/s", wdText, "Disabled"
                AddPCLproperty "aDecel", Format(.aDeceleration, "standard") & " mph/s", wdText, "Disabled"
                AddPCLproperty "aMR", .aManeuverability, wdText, "Disabled"
               
                AddPCLproperty "aSR", .aStability, wdText, "Disabled"
                AddPCLproperty "TakeOff Run (yrds)", .aTakeOffRun, wdText, "Disabled"
                AddPCLproperty "Landing Run (yrds)", .aLandingRun, wdText, "Disabled"
                AddPCLproperty "Advisory", .Advisory, wdText, "Disabled"
                
                
            Case PERFORMANCEHOVER
                AddPCLproperty "Hovercraft Performance", "", wdText, PROPERTY_HEADER
                AddPCLproperty "Hover Alt", .hHoverAltitude & " feet", wdText, "Disabled"
                AddPCLproperty "Thrust", Format(.hMotiveThrust, "standard") & " lbs", wdText, "Disabled"
                AddPCLproperty "Static Lift", Format(.staticlift, "standard") & " lbs", wdText, "Disabled"
                AddPCLproperty "Speed", Format(.hTopSpeed, "standard") & " mph", wdText, "Disabled"
                AddPCLproperty "Drag", .hDrag, wdText, "Disabled"
                AddPCLproperty "hAccel", Format(.hAcceleration, "standard") & " mph/s", wdText, "Disabled"
                AddPCLproperty "hDecel", Format(.hDeceleration, "standard") & " mph/s", wdText, "Disabled"
                AddPCLproperty "hSR", .hstability, wdText, "Disabled"
                AddPCLproperty "hMR", .hmaneuverability & " g", wdText, "Disabled"
                AddPCLproperty "Advisory", .Advisory, wdText, "Disabled"
                
                
            Case PERFORMANCEMAGLEV
                AddPCLproperty "Mag-Lev Performance", "", wdText, PROPERTY_HEADER
                AddPCLproperty "mThrust", Format(.mlMotiveThrust, "standard") & " lbs", wdText, "Disabled"
                'todo: StaticLift?  'AddPCLproperty "Static Lift", Format(.mlstaticlift, "standard") & " lbs", wdText, "Disabled"
                AddPCLproperty "mSpeed", Format(.mlTopSpeed, "standard") & " mph", wdText, "Disabled"
                AddPCLproperty "Stall Speed", Format(.mlStallSpeed, "standard") & " mph", wdText, "Disabled"
                AddPCLproperty "mDrag", .mlDrag, wdText, "Disabled"
                AddPCLproperty "mAccel", Format(.mlAcceleration, "standard") & " mph/s", wdText, "Disabled"
                AddPCLproperty "mDecel", Format(.mlDeceleration, "standard") & " mph/s", wdText, "Disabled"
                AddPCLproperty "mSR", .mlStability, wdText, "Disabled"
                AddPCLproperty "mMR", .mlManeuverability, wdText, "Disabled"
                AddPCLproperty "Advisory", .Advisory, wdText, "Disabled"
                
                
            Case PERFORMANCEWATER
                AddPCLproperty "Water Performance", "", wdText, PROPERTY_HEADER
                AddPCLproperty "wThrust", Format(.wTotalAquaticThrust, "standard") & " lbs", wdText, "Disabled"
                AddPCLproperty "wDrag", Format(.wHydroDrag, "standard"), wdText, "Disabled"
                AddPCLproperty "wSpeed", Format(.wTopSpeed, "standard") & " mph", wdText, "Disabled"
                AddPCLproperty "Hydro Speed", Format(.wHydrofoilSpeed, "standard") & " mph", wdText, "Disabled"
                AddPCLproperty "Planing Speed", Format(.wPlaningSpeed, "standard") & " mph", wdText, "Disabled"
                AddPCLproperty "wAccel", Format(.wAcceleration, "standard") & " mph/s", wdText, "Disabled"
                AddPCLproperty "wDecel", Format(.wDeceleration, "standard") & " mph/s", wdText, "Disabled"
                AddPCLproperty "Incr Decel", Format(.wIDeceleration, "standard") & " mph/s", wdText, "Disabled"
                AddPCLproperty "wSR", .wStability & "  " & "wMR: " & .wManeuverability, wdText, "Disabled"
                'AddPCLproperty  "wMR",  .wManeuverability, wdText, "Disabled"
                AddPCLproperty "wDraft", Format(.wDraft, "standard") & " feet", wdText, "Disabled"
                AddPCLproperty "Advisory", .Advisory, wdText, "Disabled"
                
                
            Case PERFORMANCESUB
                AddPCLproperty "Submerged Performance", "", wdText, PROPERTY_HEADER
                AddPCLproperty "suThrust", Format(.suTotalAquaticThrust, "standard") & " lbs", wdText, "Disabled"
                AddPCLproperty "suDrag", .suHydroDrag, wdText, "Disabled"
                AddPCLproperty "suSpeed", Format(.suTopSpeed, "standard") & " mph", wdText, "Disabled"
                AddPCLproperty "suAccel", Format(.suAcceleration, "standard") & " mph/s", wdText, "Disabled"
                AddPCLproperty "suDecel", Format(.suDeceleration, "standard") & " mph/s", wdText, "Disabled"
                AddPCLproperty "Incr Decel", Format(.suIDeceleration, "standard") & " mph/s", wdText, "Disabled"
                AddPCLproperty "suSR", .suStability, wdText, "Disabled"
                AddPCLproperty "suMR", .suManeuverability, wdText, "Disabled"
                AddPCLproperty "Draft", Format(.suDraft, "standard") & " feet", wdText, "Disabled"
                If .suCrushDepth = -1 Then
                    AddPCLproperty "Crush Depth", "No Crush Depth", wdText, "Disabled"
                Else
                    AddPCLproperty "Crush Depth", Format(.suCrushDepth, "standard") & " yards", wdText, "Disabled"
                End If
                AddPCLproperty "Advisory", .Advisory, wdText, "Disabled"
                
                
           Case PERFORMANCESPACE
                AddPCLproperty "Space Performance", "", wdText, PROPERTY_HEADER
                AddPCLproperty "Thrust", Format(.sMotiveThrust, "standard") & " lbs", wdText, "Disabled"
                'todo: make sure accel is displaying at least 4 digits for space craft
                'which have very slow accel but eventually build up to very fast speeds.
                AddPCLproperty "sAccel", Format(.sAccelerationG, "###,###,###.####") & " g", wdText, "Disabled"
                AddPCLproperty "sAccel", Format(.sAccelerationMPH, "standard") & " mph/s", wdText, "Disabled"
                AddPCLproperty "Turn Around", Format(.sTurnAroundTime, "standard") & " secs", wdText, "Disabled"
                AddPCLproperty "sMR", Format(.sManeuverability, "standard"), wdText, "Disabled"
                AddPCLproperty "Hyper", Format(.sHyperSpeed, "standard") & " parsecs", wdText, "Disabled"
                AddPCLproperty "Warp", Format(.sWarpSpeed, "standard") & " parsecs", wdText, "Disabled"
                AddPCLproperty "Jump?", .sJumpDriveable, wdText, "Disabled"
                AddPCLproperty "Teleport?", .sTeleportationDriveable, wdText, "Disabled"
                AddPCLproperty "Advisory", .Advisory, wdText, "Disabled"
        End Select
End With

End Sub

Private Sub ShowPropsForGroupComponent(ByVal component As Integer, ByVal Key As String)
    
    With m_oCurrentVeh.Components(Key)
        AddPCLproperty "Settings", "", wdText, "Disabled"
    End With
    
End Sub

Private Sub ShowPropsForSimpleCustom(ByVal component As Integer, ByVal Key As String)
    Select Case component
    
        Case SimpleCustom
            With m_oCurrentVeh.Components(Key)
                AddPCLproperty "Settings", "", wdText, "Disabled"
                AddPCLproperty "Tech level", .TL, wdList, "TL", 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
                AddPCLproperty "User Cost", .UserCost, wdDouble, "UserCost"
                AddPCLproperty "User Weight", .UserWeight, wdDouble, "UserWeight"
                AddPCLproperty "User Volume", .UserVolume, wdDouble, "UserVolume"
                AddPCLproperty "Power Consumption", .PowerReqt, wdDouble, "PowerReqt"
                AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
                AddPCLproperty "DR", .dr, wdNumber, "DR"
                AddPCLproperty "Statistics", "", wdText, "Disabled"
                AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
                AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
                AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
                AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
                AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
                
            End With
    End Select
    
End Sub

Public Sub ShowPropsForWeaponLink(ByRef sKey As String)


    With m_oCurrentVeh.WeaponProfiles(sKey)
        AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
    End With
End Sub
Private Sub ShowPropsForWeaponry1(ByVal component As Integer, ByVal Key As String)
Dim listarray() As String

With m_oCurrentVeh.Components(Key)

' Fill the window with properties for the correct Collection item
Select Case component

    Case BlueGreenLaser, RainbowLaser, Laser, UVLaser, IRLaser, Disruptor, ChargedParticleBeam, NeutralParticleBeam, _
         Flamer, Screamer, Stunner, ParalysisBeam, XRayLaser, FusionBeam, _
         GravityBeam, AntiparticleBeam, Graser, Disintegrator, Displacer, BeamedPowerTransmitter, _
         MilitaryParalysisBeam
       AddPCLproperty "Settings", "", wdText, "Disabled"
       AddPCLproperty "Tech level", .TL, wdList, "TL", 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
       AddPCLproperty "Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)"
       AddPCLproperty "Mounting", .Mount, wdList, "Mount", "normal", "concealed"
       AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
       If (component = BlueGreenLaser) Or (component = RainbowLaser) Or (component = Disintegrator) Or (component = Flamer) Or (component = Laser) Then
            AddPCLproperty "Energy Drill", .EnergyDrill, wdBool, "EnergyDrill"
       End If
       AddPCLproperty "Beam Output", .BeamOutput, wdDouble, "BeamOutput"
       AddPCLproperty "Cyclic Rate", .rof, wdList, "rof", "1/14400", "1/7200", "1/4800", "1/3600", "1/2400", "1/1200", "1/600", "1/300", "1/150", "1/60", "1/30", "1/15", "1/8", "1/4", "1/2", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20"
       AddPCLproperty "Range", .Range, wdList, "Range", "close", "normal", "long", "very long", "extreme"
       AddPCLproperty "Power Cells", .PowerCellType, wdList, "PowerCellType", "none", "C cells", "rC cell", "D cells", "rD cells", "E cells", "rE cells"
       AddPCLproperty "# Power Cells", .PowerCellQuantity, wdNumber, "PowerCellQuantity"
       AddPCLproperty "FTL Beam?", .FTL, wdBool, "FTL"
       AddPCLproperty "Compact", .Compact, wdBool, "Compact"
       AddPCLproperty "Reputation for Quality", .Reliable, wdBool, "Reliable"
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
       AddPCLproperty "DR", .dr, wdNumber, "DR"
       AddPCLproperty "Statistics", "", wdText, "Disabled"
       AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
       AddPCLproperty "Type Damage", .TypeDamage, wdText, "Disabled"
        'note: for displacers its radius of effect and not damage
        'note: for paralysis beams its HT penalty and not damage
        'note: for stunners its HT penalty also and not damage
       If (component = Stunner) Or (component = ParalysisBeam) Or (component = MilitaryParalysisBeam) Then
            AddPCLproperty "HT penalty", .Damage, wdText, "Disabled"
       Else
           AddPCLproperty "Damage", .Damage, wdText, "Disabled"
        End If
        'note: militaryparalysis,paralysis, disintegrators and displacers have no half damage
       AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
       AddPCLproperty "Vacuum 1/2 Damage (yards)", .VacuumHalfDamage, wdText, "Disabled"
       
        If component = Stunner Then
           AddPCLproperty "Max Range at HT6- (yards)", .MaxRange, wdText, "Disabled"
           AddPCLproperty "Max Range at HT7+ (yards)", .MaxRange2, wdText, "Disabled"
           AddPCLproperty "Vacuum Max Range at HT6- (yards)", .VacuumMaxRange, wdText, "Disabled"
           AddPCLproperty "Vacuum Max Range at HT7+ (yards)", .VacuumMaxRange2, wdText, "Disabled"
        Else
           AddPCLproperty "Max Range (yards)", .MaxRange, wdText, "Disabled"
           AddPCLproperty "Vacuum Max Range (yards)", .VacuumMaxRange, wdText, "Disabled"
        End If
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
       AddPCLproperty "Snap Shot", .SnapShot, wdText, "Disabled"
       AddPCLproperty "Rate of Fire", .rof, wdText, "Disabled"
       AddPCLproperty "# of Shots", .Shots, wdText, "Disabled"
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
       
        
     Case UnGuidedMissile, UnGuidedTorpedo
       AddPCLproperty "Settings", "", wdText, "Disabled"
       AddPCLproperty "Tech level", .TL, wdList, "TL", 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
       AddPCLproperty "Diameter", .Diameter, wdDouble, "Diameter"
       listarray = .FillAmmunitionList
       AddPCLproperty "Warhead", .WarHead, wdList, "Warhead", listarray
       AddPCLproperty "Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge"
       AddPCLproperty "# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"
        If .SpaceMissile = False Then
           AddPCLproperty "Speed (yds per sec)", .Speed, wdDouble, "Speed"
        Else
           AddPCLproperty "G's", .Speed, wdDouble, "Speed"
        End If
       AddPCLproperty "Motor Weight", .MotorWeight, wdDouble, "MotorWeight"
       AddPCLproperty "Stealth", .stealth, wdBool, "Stealth"
        'this option only available for relevant missile types
       AddPCLproperty "Space Fairing?", .SpaceMissile, wdBool, "SpaceMissile"
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
       AddPCLproperty "DR", .dr, wdNumber, "DR"
       AddPCLproperty "Statistics", "", wdText, "Disabled"
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
       If .BurstRadius <> -1 Then
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
        If .TypeDamage2 <> "none" Then
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
       AddPCLproperty "Endurance (seconds)", .Endurance, wdText, "Disabled"
        'only unguided missiles have 1/2 damage
        If component = UnGuidedMissile Then
           AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
        End If
       AddPCLproperty "Max Range (yards)", .MaxRange, wdText, "Disabled"
       AddPCLproperty "Min Range (yards)", .MinRange, wdText, "Disabled"
       AddPCLproperty "Motor Cost", "$" & Format(.MotorCost, "standard"), wdText, "Disabled"
       AddPCLproperty "Warhead Cost", "$" & Format(.WarheadCost, "standard"), wdText, "Disabled"
       AddPCLproperty "Warhead Weight", Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Payload Cost", "$" & Format(.PayloadCost, "standard"), wdText, "Disabled"
       AddPCLproperty "Payload Weight", Format(.PayloadWeight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Total Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
       AddPCLproperty "Total Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
        
        
    Case GuidedMissile, GuidedTorpedo
       AddPCLproperty "Settings", "", wdText, "Disabled"
       AddPCLproperty "Tech level", .TL, wdList, "TL", 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
       listarray = .FillGuidanceList
       AddPCLproperty "Guidance System", .GuidanceSystem, wdList, "GuidanceSystem", listarray
       listarray = .FillTerminalGuidanceList
       AddPCLproperty "Terminal Guidance", .BrilliantGuidanceSystem, wdList, "BrilliantGuidanceSystem", listarray
       AddPCLproperty "Cheap Guidance System", .CheapGuidance, wdBool, "CheapGuidance"
       AddPCLproperty "Compact Guidance System", .Compact, wdBool, "Compact"
       AddPCLproperty "Mid-Course Update", .MidCourseUpdate, wdBool, "MidCourseUpdate"
       AddPCLproperty "Pop-Up", .PopUp, wdBool, "Popup"
       AddPCLproperty "Diameter", .Diameter, wdDouble, "Diameter"
       listarray = .FillAmmunitionList
       AddPCLproperty "Warhead", .WarHead, wdList, "Warhead", listarray
       AddPCLproperty "Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge"
       AddPCLproperty "# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"
       AddPCLproperty "Skill Bonus", .SkillBonus, wdNumber, "SkillBonus"
        'this option only available for relevant missile /torp types
        If .SpaceMissile = False Then
           AddPCLproperty "Speed (yds per sec)", .Speed, wdDouble, "Speed"
        Else
           AddPCLproperty "G's", .Speed, wdDouble, "Speed"
        End If
       AddPCLproperty "Motor Weight", .MotorWeight, wdDouble, "MotorWeight"
       AddPCLproperty "Stealth", .stealth, wdBool, "Stealth"
        'this option only available for relevant missile types
       AddPCLproperty "Space Fairing?", .SpaceMissile, wdBool, "SpaceMissile"
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
       AddPCLproperty "DR", .dr, wdNumber, "DR"
       AddPCLproperty "Statistics", "", wdText, "Disabled"
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
       If .BurstRadius <> -1 Then
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
        If .TypeDamage2 <> "none" Then
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
       AddPCLproperty "Skill", .Skill, wdText, "Disabled"
       AddPCLproperty "Endurance (seconds)", .Endurance, wdText, "Disabled"
       AddPCLproperty "Max Range (yards)", .MaxRange, wdText, "Disabled"
       AddPCLproperty "Min Range (yards)", .MinRange, wdText, "Disabled"
       AddPCLproperty "Motor Cost", "$" & Format(.MotorCost, "standard"), wdText, "Disabled"
       AddPCLproperty "Warhead Cost", "$" & Format(.WarheadCost, "standard"), wdText, "Disabled"
       AddPCLproperty "Warhead Weight", Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Guidance System Cost", "$" & Format(.GuidanceCost, "standard"), wdText, "Disabled"
       AddPCLproperty "Guidance System Weight", Format(.GuidanceWeight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Payload Cost", "$" & Format(.PayloadCost, "standard"), wdText, "Disabled"
       AddPCLproperty "Payload Weight", Format(.PayloadWeight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Total Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
       AddPCLproperty "Total Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
       
    Case IronBomb, SelfDestructSystem
       AddPCLproperty "Settings", "", wdText, "Disabled"
       AddPCLproperty "Tech level", .TL, wdList, "TL", 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
       AddPCLproperty "Diameter", .Diameter, wdDouble, "Diameter"
       listarray = .FillAmmunitionList
       AddPCLproperty "Warhead", .WarHead, wdList, "Warhead", listarray
       AddPCLproperty "Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge"
       AddPCLproperty "# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"
       'note: these dont have a speed do they?
       'AddPCLproperty "Speed (yds per sec)", .Speed, wdDouble, "Speed"
       AddPCLproperty "Stealth", .stealth, wdBool, "Stealth"
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
       AddPCLproperty "DR", .dr, wdNumber, "DR"
       AddPCLproperty "Statistics", "", wdText, "Disabled"
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
       If .BurstRadius <> -1 Then
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
        If .TypeDamage2 <> "none" Then
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
       AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
       AddPCLproperty "Min Range (yards)", .MinRange, wdText, "Disabled"
       AddPCLproperty "Warhead Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
       AddPCLproperty "Warhead Weight", Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Cost", "$" & Format(.PayloadCost, "standard"), wdText, "Disabled"
       AddPCLproperty "Weight", Format(.PayloadWeight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
        
    
    Case RetardedBomb
       AddPCLproperty "Settings", "", wdText, "Disabled"
       AddPCLproperty "Tech level", .TL, wdList, "TL", 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
       AddPCLproperty "Diameter", .Diameter, wdDouble, "Diameter"
       listarray = .FillAmmunitionList
       AddPCLproperty "Warhead", .WarHead, wdList, "Warhead", listarray
       AddPCLproperty "Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge"
       AddPCLproperty "# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"
       AddPCLproperty "Speed (yds per sec)", .Speed, wdDouble, "Speed"
       AddPCLproperty "Stealth", .stealth, wdBool, "Stealth"
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
       AddPCLproperty "DR", .dr, wdNumber, "DR"
       AddPCLproperty "Statistics", "", wdText, "Disabled"
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
       If .BurstRadius <> -1 Then
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
        If .TypeDamage2 <> "none" Then
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
       AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
       AddPCLproperty "Min Range (yards)", .MinRange, wdText, "Disabled"
       AddPCLproperty "Warhead Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
       AddPCLproperty "Payload Cost", "$" & Format(.PayloadCost, "standard"), wdText, "Disabled"
       AddPCLproperty "Warhead Weight", Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Payload Weight", Format(.PayloadWeight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
        
    
    Case SmartBomb
       AddPCLproperty "Settings", "", wdText, "Disabled"
       AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
       listarray = .FillGuidanceList
       AddPCLproperty "Guidance System", .GuidanceSystem, wdList, "GuidanceSystem", listarray
       AddPCLproperty "Cheap Guidance System", .CheapGuidance, wdBool, "CheapGuidance"
       AddPCLproperty "Compact Guidance System", .Compact, wdBool, "Compact"
       AddPCLproperty "Diameter", .Diameter, wdDouble, "Diameter"
       listarray = .FillAmmunitionList
       AddPCLproperty "Warhead", .WarHead, wdList, "Warhead", listarray
       AddPCLproperty "Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge"
       AddPCLproperty "# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"
       AddPCLproperty "Skill Bonus", .SkillBonus, wdNumber, "SkillBonus"
       AddPCLproperty "Speed (yds per sec)", .Speed, wdDouble, "Speed"
       AddPCLproperty "Stealth", .stealth, wdBool, "Stealth"
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
       AddPCLproperty "DR", .dr, wdNumber, "DR"
       AddPCLproperty "Statistics", "", wdText, "Disabled"
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
       If .BurstRadius <> -1 Then
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
        If .TypeDamage2 <> "none" Then
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
       AddPCLproperty "Skill", .Skill, wdText, "Disabled"
       AddPCLproperty "Min Range (yards)", .MinRange, wdText, "Disabled"
       AddPCLproperty "Warhead Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
       AddPCLproperty "Guidance System Cost", "$" & Format(.GuidanceCost, "standard"), wdText, "Disabled"
       AddPCLproperty "Payload Cost", "$" & Format(.PayloadCost, "standard"), wdText, "Disabled"
       AddPCLproperty "Warhead Weight", Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Guidance System Weight", Format(.GuidanceWeight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Payload Weight", Format(.PayloadWeight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
        
    
    Case ContactMine
       AddPCLproperty "Settings", "", wdText, "Disabled"
       AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
       AddPCLproperty "Diameter", .Diameter, wdDouble, "Diameter"
       listarray = .FillAmmunitionList
       AddPCLproperty "Warhead", .WarHead, wdList, "Warhead", listarray
       AddPCLproperty "Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge"
       AddPCLproperty "# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
       AddPCLproperty "DR", .dr, wdNumber, "DR"
       AddPCLproperty "Statistics", "", wdText, "Disabled"
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
       If .BurstRadius <> -1 Then
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
        If .TypeDamage2 <> "none" Then
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
       AddPCLproperty "Warhead Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
       AddPCLproperty "Warhead Weight", Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
        
    Case ProximityMine
       AddPCLproperty "Settings", "", wdText, "Disabled"
       AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
       listarray = .FillGuidanceList
       AddPCLproperty "Guidance System", .GuidanceSystem, wdList, "GuidanceSystem", listarray
       AddPCLproperty "Cheap Guidance System", .CheapGuidance, wdBool, "CheapGuidance"
       AddPCLproperty "Compact Guidance System", .Compact, wdBool, "Compact"
       AddPCLproperty "Diameter", .Diameter, wdDouble, "Diameter"
       listarray = .FillAmmunitionList
       AddPCLproperty "Warhead", .WarHead, wdList, "Warhead", listarray
       AddPCLproperty "Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge"
       AddPCLproperty "# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"
       AddPCLproperty "Skill Bonus", .SkillBonus, wdNumber, "SkillBonus"
       AddPCLproperty "Stealth", .stealth, wdBool, "Stealth"
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
       AddPCLproperty "DR", .dr, wdNumber, "DR"
       AddPCLproperty "Statistics", "", wdText, "Disabled"
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
       If .BurstRadius <> -1 Then
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
        If .TypeDamage2 <> "none" Then
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
       AddPCLproperty "Skill", .Skill, wdText, "Disabled"
       AddPCLproperty "Warhead Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
       AddPCLproperty "Guidance System Cost", "$" & Format(.GuidanceCost, "standard"), wdText, "Disabled"
       AddPCLproperty "Payload Cost", "$" & Format(.PayloadCost, "standard"), wdText, "Disabled"
       AddPCLproperty "Warhead Weight", Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Guidance System Weight", Format(.GuidanceWeight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Payload Weight", Format(.PayloadWeight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
        
    
    Case PressureTriggerMine
       AddPCLproperty "Settings", "", wdText, "Disabled"
       AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
       AddPCLproperty "Diameter", .Diameter, wdDouble, "Diameter"
       AddPCLproperty "Detonation Weight", .DetonationWeight, wdDouble, "DetonationWeight"
       listarray = .FillAmmunitionList
       AddPCLproperty "Warhead", .WarHead, wdList, "Warhead", listarray
       AddPCLproperty "Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge"
       AddPCLproperty "# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
       AddPCLproperty "DR", .dr, wdNumber, "DR"
       AddPCLproperty "Statistics", "", wdText, "Disabled"
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
       If .BurstRadius <> -1 Then
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
        If .TypeDamage2 <> "none" Then
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
       AddPCLproperty "Warhead Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
       AddPCLproperty "Warhead Weight", Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
        
        
    Case CommandTriggerMine, SmartTriggerMine
       AddPCLproperty "Settings", "", wdText, "Disabled"
       AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
       AddPCLproperty "Diameter", .Diameter, wdDouble, "Diameter"
       AddPCLproperty "Parachute Mine?", .Parachute, wdBool, "Parachute"
       listarray = .FillAmmunitionList
       AddPCLproperty "Warhead", .WarHead, wdList, "Warhead", listarray
       AddPCLproperty "Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge"
       AddPCLproperty "# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
       AddPCLproperty "DR", .dr, wdNumber, "DR"
       AddPCLproperty "Statistics", "", wdText, "Disabled"
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
       If .BurstRadius <> -1 Then
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
        If .TypeDamage2 <> "none" Then
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
       AddPCLproperty "Warhead Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
       AddPCLproperty "Warhead Weight", Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
        
    
    Case WaterCannon, FlameThrower
       AddPCLproperty "Settings", "", wdText, "Disabled"
       AddPCLproperty "Tech level", .TL, wdList, "TL", 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
       AddPCLproperty "Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)"
       AddPCLproperty "Mounting", .Mount, wdList, "Mount", "normal", "concealed"
       AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
       AddPCLproperty "Style", .Style, wdList, "Style", "light", "medium", "heavy"
       If component = WaterCannon Then
           AddPCLproperty "Type of Ammo", .Ammunitiontype, wdList, "AmmunitionType", "water", "acid", "foam"
       End If
       AddPCLproperty "# of Shots", .Shots, wdNumber, "Shots"
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
       AddPCLproperty "DR", .dr, wdNumber, "DR"
       AddPCLproperty "Statistics", "", wdText, "Disabled"
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
       AddPCLproperty "Type Damage", .TypeDamage, wdText, "Disabled"
       AddPCLproperty "Damage", .Damage, wdText, "Disabled"
       AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
       AddPCLproperty "Max Range (yards)", .MaxRange, wdText, "Disabled"
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
       AddPCLproperty "Snap Shot", .SnapShot, wdText, "Disabled"
       AddPCLproperty "Rate of Fire", .rof, wdText, "Disabled"
       AddPCLproperty "# of Shots", .Shots, wdText, "Disabled"
       AddPCLproperty "Cost Per Shot", .CPS, wdText, "Disabled"
       AddPCLproperty "Weight Per Shot", .WPS, wdText, "Disabled"
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Loaders", .Loaders, wdText, "Disabled"
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
        
    Case RevolverLauncher, DisposableLauncher, MuzzleloadingLauncher, BreechloadingLauncher, _
        ManualRepeaterLauncher, SlowAutoLoaderLauncher, FastAutoLoaderLauncher, _
         lightAutomaticLauncher, HeavyAutomaticLauncher
       AddPCLproperty "Settings", "", wdText, "Disabled"
       AddPCLproperty "Tech level", .TL, wdList, "TL", 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
       AddPCLproperty "Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)"
       AddPCLproperty "Mounting", .Mount, wdList, "Mount", "normal", "concealed"
       AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
       AddPCLproperty "Diameter", .Diameter, wdDouble, "Diameter"
       AddPCLproperty "Maximum Load (lbs)", .MaxLoad, wdDouble, "MaxLoad"
       Select Case component
        Case RevolverLauncher, DisposableLauncher, MuzzleloadingLauncher, BreechloadingLauncher
            AddPCLproperty "# of Tubes", .Cylinders, wdNumber, "Cylinders"
       End Select
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
       AddPCLproperty "DR", .dr, wdNumber, "DR"
       AddPCLproperty "Statistics", "", wdText, "Disabled"
       AddPCLproperty "Snap Shot", .SnapShot, wdText, "Disabled"
       AddPCLproperty "Rate of Fire", .rof, wdText, "Disabled"
       AddPCLproperty "# of Shots", .Shots, wdText, "Disabled"
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Loaders", .Loaders, wdText, "Disabled"
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
         
End Select
End With

End Sub

Private Sub ShowPropsForWeaponry2(ByVal component As Integer, ByVal Key As String)
Dim listarray() As String

With m_oCurrentVeh.Components(Key)

' Fill the window with properties for the correct Collection item
Select Case component

    Case StoneThrower, BoltThrower
       AddPCLproperty "Settings", "", wdText, "Disabled"
       AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
       AddPCLproperty "Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)"
       AddPCLproperty "Mounting", .Mount, wdList, "Mount", "normal", "concealed"
       AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
       If component = StoneThrower Then
            AddPCLproperty "Mechanism", .Mechanism, wdList, "Mechanism", "spring-powered", "torsion-powered", "counterweight"
       Else
            AddPCLproperty "Mechanism", .Mechanism, wdList, "Mechanism", "spring-powered", "torsion-powered"
       End If
       AddPCLproperty "Strength", .Strength, wdNumber, "Strength"
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
       AddPCLproperty "DR", .dr, wdNumber, "DR"
       AddPCLproperty "Statistics", "", wdText, "Disabled"
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
       AddPCLproperty "Type Damage", .TypeDamage, wdText, "Disabled"
       AddPCLproperty "Damage", .Damage, wdText, "Disabled"
       AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
       AddPCLproperty "Max Range (yards)", .MaxRange, wdText, "Disabled"
       AddPCLproperty "Min Range (yards)", .MinRange, wdText, "Disabled"
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
       AddPCLproperty "Snap Shot", .SnapShot, wdText, "Disabled"
       AddPCLproperty "Rate of Fire", .rof, wdText, "Disabled"
       AddPCLproperty "# of Shots", .Shots, wdText, "Disabled"
       AddPCLproperty "Reqt. Loaders", .Loaders, wdText, "Disabled"
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
       AddPCLproperty "Cost Per Shot", "$" & Format(.CPS, "standard"), wdText, "Disabled"
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Weight Per Shot", Format(.WPS, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Volume Per Shot", Format(.VPS, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
       
    Case RepeatingBoltThrower
       AddPCLproperty "Settings", "", wdText, "Disabled"
       AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
       AddPCLproperty "Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)"
       AddPCLproperty "Mounting", .Mount, wdList, "Mount", "normal", "concealed"
       AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
       AddPCLproperty "Mechanism", .Mechanism, wdList, "Mechanism", "spring-powered", "torsion-powered"
       AddPCLproperty "Strength", .Strength, wdNumber, "Strength"
       AddPCLproperty "Magazine Capacity", .MagazineCapacity, wdNumber, "MagazineCapacity"
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
       AddPCLproperty "DR", .dr, wdNumber, "DR"
       AddPCLproperty "Statistics", "", wdText, "Disabled"
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
       AddPCLproperty "Type Damage", .TypeDamage, wdText, "Disabled"
       AddPCLproperty "Damage", .Damage, wdText, "Disabled"
       AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
       AddPCLproperty "Max Range (yards)", .MaxRange, wdText, "Disabled"
       AddPCLproperty "Min Range (yards)", .MinRange, wdText, "Disabled"
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
       AddPCLproperty "Snap Shot", .SnapShot, wdText, "Disabled"
       AddPCLproperty "Rate of Fire", .rof, wdText, "Disabled"
       AddPCLproperty "# of Shots", .Shots, wdText, "Disabled"
       AddPCLproperty "Reqt. Loaders", .Loaders, wdText, "Disabled"
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
       AddPCLproperty "Cost Per Shot", "$" & Format(.CPS, "standard"), wdText, "Disabled"
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Weight Per Shot", Format(.WPS, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Volume Per Shot", Format(.VPS, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
        
    
    
    Case MuzzleLoader, BreechLoader
       AddPCLproperty "Settings", "", wdText, "Disabled"
       AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
       AddPCLproperty "Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)"
       AddPCLproperty "Mounting", .Mount, wdList, "Mount", "normal", "concealed"
       AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
       AddPCLproperty "Bore Size", .BoreSize, wdDouble, "BoreSize"
       AddPCLproperty "Technology", .technology, wdList, "Technology", "conventional smoothbore", "conventional rifled", "electromag", "gravitic"
       listarray = .FillAmmunitionList
       AddPCLproperty "Ammunition", .Ammunitiontype, wdList, "AmmunitionType", listarray
       AddPCLproperty "Power Option", .PowerOption, wdList, "PowerOption", "normal", "low-powered", "extra-low-powered"
       AddPCLproperty "Recoiless", .Recoiless, wdBool, "Recoiless"
       AddPCLproperty "Barrel Length", .Barrel, wdList, "Barrel", "extremely short", "very short", "short", "medium", "long", "very long", "extremely long"
       AddPCLproperty "# of Fixed Barrels", .Cylinders, wdList, "Cylinders", "1", "2", "3", "4", "5", "6", "7"
       'advanced option not available for unconventinal weapons
       If (.technology = "conventional smoothbore") Or (.technology = "conventional rifled") Then
            AddPCLproperty "Advanced Option", .advancedoption, wdList, "AdvancedOption", "none", "plastic-cased ammunition", "caseless", "liquid propellant", "electrothermal"
       Else
            .advancedoption = "none"
       End If
       
       AddPCLproperty "Reputation for Quality", .Reliable, wdBool, "Reliable"
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
       AddPCLproperty "DR", .dr, wdNumber, "DR"
       AddPCLproperty "Statistics", "", wdText, "Disabled"
        If (component = ElectricGatling) Or (.technology = "electromag") Or (.Electric) Then
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
        End If
        If component = MuzzleLoader Then
           AddPCLproperty "Carriage Required", .Carriage, wdText, "Disabled"
        End If
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
       If .BurstRadius <> -1 Then
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
        If .TypeDamage2 <> "none" Then
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
       AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
       AddPCLproperty "Max Range (yards)", .MaxRange, wdText, "Disabled"
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
       AddPCLproperty "Snap Shot", .SnapShot, wdText, "Disabled"
       AddPCLproperty "Rate of Fire", .sRoF, wdText, "Disabled"
       AddPCLproperty "# of Shots", .Shots, wdText, "Disabled"
       AddPCLproperty "Reqt. Loaders", .Loaders, wdText, "Disabled"
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
       AddPCLproperty "Cost Per Shot", "$" & Format(.CPS, "standard"), wdText, "Disabled"
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Weight Per Shot", Format(.WPS, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Volume Per Shot", Format(.VPS, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
        
    
    Case ManualRepeater
       AddPCLproperty "Settings", "", wdText, "Disabled"
       AddPCLproperty "Tech level", .TL, wdList, "TL", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16"
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
       AddPCLproperty "Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)"
       AddPCLproperty "Mounting", .Mount, wdList, "Mount", "normal", "concealed"
       AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
       AddPCLproperty "Bore Size", .BoreSize, wdDouble, "BoreSize"
       AddPCLproperty "Technology", .technology, wdList, "Technology", "conventional smoothbore", "conventional rifled", "electromag", "gravitic"
       listarray = .FillAmmunitionList
       AddPCLproperty "Ammunition", .Ammunitiontype, wdList, "AmmunitionType", listarray
       AddPCLproperty "Power Option", .PowerOption, wdList, "PowerOption", "normal", "low-powered", "extra-low-powered"
       AddPCLproperty "Recoiless", .Recoiless, wdBool, "Recoiless"
       AddPCLproperty "Barrel Length", .Barrel, wdList, "Barrel", "extremely short", "very short", "short", "medium", "long", "very long", "extremely long"
       'advanced option not available for unconventinal weapons
       If (.technology = "conventional smoothbore") Or (.technology = "conventional rifled") Then
            AddPCLproperty "Advanced Option", .advancedoption, wdList, "AdvancedOption", "none", "plastic-cased ammunition", "caseless", "liquid propellant", "electrothermal"
       Else
            .advancedoption = "none"
       End If
       AddPCLproperty "Box Magazine", .BoxMagazine, wdBool, "BoxMagazine"
       AddPCLproperty "Reputation for Quality", .Reliable, wdBool, "Reliable"
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
       AddPCLproperty "DR", .dr, wdNumber, "DR"
       AddPCLproperty "Statistics", "", wdText, "Disabled"
        If (component = ElectricGatling) Or (.technology = "electromag") Or (.Electric) Then
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
        End If
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
       If .BurstRadius <> -1 Then
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
        If .TypeDamage2 <> "none" Then
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
       AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
       AddPCLproperty "Max Range (yards)", .MaxRange, wdText, "Disabled"
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
       AddPCLproperty "Snap Shot", .SnapShot, wdText, "Disabled"
       AddPCLproperty "Rate of Fire", .sRoF, wdText, "Disabled"
       AddPCLproperty "# of Shots", .Shots, wdText, "Disabled"
       AddPCLproperty "Reqt. Loaders", .Loaders, wdText, "Disabled"
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
       AddPCLproperty "Cost Per Shot", "$" & Format(.CPS, "standard"), wdText, "Disabled"
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Weight Per Shot", Format(.WPS, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Volume Per Shot", Format(.VPS, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
        
        
    Case Revolver, MechanicalGatling
            'NOTE: These allow for user modifieable Rates of Fire
    'power only needs to be displayed for elec.gat.
       AddPCLproperty "Settings", "", wdText, "Disabled"
       AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
       AddPCLproperty "Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)"
       AddPCLproperty "Mounting", .Mount, wdList, "Mount", "normal", "concealed"
       AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
       AddPCLproperty "Bore Size", .BoreSize, wdDouble, "BoreSize"
       AddPCLproperty "Technology", .technology, wdList, "Technology", "conventional smoothbore", "conventional rifled", "electromag", "gravitic"
       listarray = .FillAmmunitionList
       AddPCLproperty "Ammunition", .Ammunitiontype, wdList, "AmmunitionType", listarray
       If component = MechanicalGatling Then
           AddPCLproperty "Operator DX + Skill", .DXPlusSkill, wdDouble, "DXPlusSkill"
       End If
       AddPCLproperty "Power Option", .PowerOption, wdList, "PowerOption", "normal", "low-powered", "extra-low-powered"
       AddPCLproperty "Recoiless", .Recoiless, wdBool, "Recoiless"
       AddPCLproperty "Barrel Length", .Barrel, wdList, "Barrel", "extremely short", "very short", "short", "medium", "long", "very long", "extremely long"
        If component = Revolver Then
           AddPCLproperty "# of Cylinders", .Cylinders, wdList, "Cylinders", "3", "4", "5", "6", "7"
        Else
           AddPCLproperty "# of Barrels", .Cylinders, wdList, "Cylinders", "3", "4", "5", "6", "7"
        End If
       'advanced option not available for unconventinal weapons
       If (.technology = "conventional smoothbore") Or (.technology = "conventional rifled") Then
            AddPCLproperty "Advanced Option", .advancedoption, wdList, "AdvancedOption", "none", "plastic-cased ammunition", "caseless", "liquid propellant", "electrothermal"
       Else
            .advancedoption = "none"
       End If
       AddPCLproperty "Reputation for Quality", .Reliable, wdBool, "Reliable"
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
       AddPCLproperty "DR", .dr, wdNumber, "DR"
       AddPCLproperty "Statistics", "", wdText, "Disabled"
        If (component = ElectricGatling) Or (.technology = "electromag") Or (.Electric) Then
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
        End If
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
       If .BurstRadius <> -1 Then
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
        If .TypeDamage2 <> "none" Then
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
       AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
       AddPCLproperty "Max Range (yards)", .MaxRange, wdText, "Disabled"
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
       AddPCLproperty "Snap Shot", .SnapShot, wdText, "Disabled"
       AddPCLproperty "Rate of Fire", .sRoF, wdText, "Disabled"
       AddPCLproperty "# of Shots", .Shots, wdText, "Disabled"
       AddPCLproperty "Reqt. Loaders", .Loaders, wdText, "Disabled"
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
       AddPCLproperty "Cost Per Shot", "$" & Format(.CPS, "standard"), wdText, "Disabled"
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Weight Per Shot", Format(.WPS, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Volume Per Shot", Format(.VPS, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
    
    Case ElectricGatling
    'NOTE: This allows for user modifieable Rate of Fire
       AddPCLproperty "Settings", "", wdText, "Disabled"
       AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
       AddPCLproperty "Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)"
       AddPCLproperty "Mounting", .Mount, wdList, "Mount", "normal", "concealed"
       AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
       AddPCLproperty "Bore Size", .BoreSize, wdDouble, "BoreSize"
       AddPCLproperty "Technology", .technology, wdList, "Technology", "conventional smoothbore", "conventional rifled", "electromag", "gravitic"
       listarray = .FillAmmunitionList
       AddPCLproperty "Ammunition", .Ammunitiontype, wdList, "AmmunitionType", listarray
       listarray = .FillRoFList
       AddPCLproperty "Rate of Fire", .dRoF, wdList, "dRoF", listarray
       AddPCLproperty "Power Option", .PowerOption, wdList, "PowerOption", "normal", "low-powered", "extra-low-powered"
       AddPCLproperty "Recoiless", .Recoiless, wdBool, "Recoiless"
       AddPCLproperty "Barrel Length", .Barrel, wdList, "Barrel", "extremely short", "very short", "short", "medium", "long", "very long", "extremely long"
        If component = Revolver Then
           AddPCLproperty "# of Cylinders", .Cylinders, wdList, "Cylinders", "3", "4", "5", "6", "7"
        Else
           AddPCLproperty "# of Barrels", .Cylinders, wdList, "Cylinders", "3", "4", "5", "6", "7"
        End If
       'advanced option not available for unconventinal weapons
       If (.technology = "conventional smoothbore") Or (.technology = "conventional rifled") Then
            AddPCLproperty "Advanced Option", .advancedoption, wdList, "AdvancedOption", "none", "plastic-cased ammunition", "caseless", "liquid propellant", "electrothermal"
       Else
            .advancedoption = "none"
       End If
       AddPCLproperty "Reputation for Quality", .Reliable, wdBool, "Reliable"
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
       AddPCLproperty "DR", .dr, wdNumber, "DR"
       AddPCLproperty "Statistics", "", wdText, "Disabled"
        If (component = ElectricGatling) Or (.technology = "electromag") Or (.Electric) Then
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
        End If
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
       If .BurstRadius <> -1 Then
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
        If .TypeDamage2 <> "none" Then
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
       AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
       AddPCLproperty "Max Range (yards)", .MaxRange, wdText, "Disabled"
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
       AddPCLproperty "Snap Shot", .SnapShot, wdText, "Disabled"
       AddPCLproperty "# of Shots", .Shots, wdText, "Disabled"
       AddPCLproperty "Reqt. Loaders", .Loaders, wdText, "Disabled"
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
       AddPCLproperty "Cost Per Shot", "$" & Format(.CPS, "standard"), wdText, "Disabled"
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Weight Per Shot", Format(.WPS, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Volume Per Shot", Format(.VPS, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
        
        
    Case SlowAutoloader, FastAutoloader
        AddPCLproperty "Settings", "", wdText, "Disabled"
       AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
       AddPCLproperty "Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)"
       AddPCLproperty "Mounting", .Mount, wdList, "Mount", "normal", "concealed"
       AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
       AddPCLproperty "Bore Size", .BoreSize, wdDouble, "BoreSize"
       AddPCLproperty "Technology", .technology, wdList, "Technology", "conventional smoothbore", "conventional rifled", "electromag", "gravitic"
       listarray = .FillAmmunitionList
       AddPCLproperty "Ammunition", .Ammunitiontype, wdList, "AmmunitionType", listarray
       AddPCLproperty "Power Option", .PowerOption, wdList, "PowerOption", "normal", "low-powered", "extra-low-powered"
       AddPCLproperty "Recoiless", .Recoiless, wdBool, "Recoiless"
       AddPCLproperty "Barrel Length", .Barrel, wdList, "Barrel", "extremely short", "very short", "short", "medium", "long", "very long", "extremely long"
       AddPCLproperty "Electric Loading", .Electric, wdBool, "Electric"
       'advanced option not available for unconventinal weapons
       If (.technology = "conventional smoothbore") Or (.technology = "conventional rifled") Then
            AddPCLproperty "Advanced Option", .advancedoption, wdList, "AdvancedOption", "none", "plastic-cased ammunition", "caseless", "liquid propellant", "electrothermal"
       Else
            .advancedoption = "none"
       End If
       AddPCLproperty "Reputation for Quality", .Reliable, wdBool, "Reliable"
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
       AddPCLproperty "DR", .dr, wdNumber, "DR"
       AddPCLproperty "Statistics", "", wdText, "Disabled"
        If (component = ElectricGatling) Or (.technology = "electromag") Or (.technology = "gravitic") Or (.Electric) Then
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
        End If
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
       If .BurstRadius <> -1 Then
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
        If .TypeDamage2 <> "none" Then
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
       AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
       AddPCLproperty "Max Range (yards)", .MaxRange, wdText, "Disabled"
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
       AddPCLproperty "Snap Shot", .SnapShot, wdText, "Disabled"
       AddPCLproperty "Rate of Fire", .sRoF, wdText, "Disabled"
       AddPCLproperty "# of Shots", .Shots, wdText, "Disabled"
       AddPCLproperty "Reqt. Loaders", .Loaders, wdText, "Disabled"
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
       AddPCLproperty "Cost Per Shot", "$" & Format(.CPS, "standard"), wdText, "Disabled"
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Weight Per Shot", Format(.WPS, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Volume Per Shot", Format(.VPS, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
       
    Case lightAutomatic, HeavyAutomatic
        'note: these allow for user edit-able Rates of Fire
       AddPCLproperty "Settings", "", wdText, "Disabled"
       AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
       AddPCLproperty "Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)"
       AddPCLproperty "Mounting", .Mount, wdList, "Mount", "normal", "concealed"
       AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
       AddPCLproperty "Bore Size", .BoreSize, wdDouble, "BoreSize"
       AddPCLproperty "Technology", .technology, wdList, "Technology", "conventional smoothbore", "conventional rifled", "electromag", "gravitic"
       listarray = .FillAmmunitionList
       AddPCLproperty "Ammunition", .Ammunitiontype, wdList, "AmmunitionType", listarray
       listarray = .FillRoFList
       AddPCLproperty "Rate of Fire", .dRoF, wdList, "dRoF", listarray
       AddPCLproperty "Power Option", .PowerOption, wdList, "PowerOption", "normal", "low-powered", "extra-low-powered"
       AddPCLproperty "Recoiless", .Recoiless, wdBool, "Recoiless"
       AddPCLproperty "Barrel Length", .Barrel, wdList, "Barrel", "extremely short", "very short", "short", "medium", "long", "very long", "extremely long"
       AddPCLproperty "Electric Loading", .Electric, wdBool, "Electric"
       'advanced option not available for unconventinal weapons
       If (.technology = "conventional smoothbore") Or (.technology = "conventional rifled") Then
            AddPCLproperty "Advanced Option", .advancedoption, wdList, "AdvancedOption", "none", "plastic-cased ammunition", "caseless", "liquid propellant", "electrothermal"
       Else
            .advancedoption = "none"
       End If
       AddPCLproperty "Reputation for Quality", .Reliable, wdBool, "Reliable"
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
       AddPCLproperty "DR", .dr, wdNumber, "DR"
       AddPCLproperty "Statistics", "", wdText, "Disabled"
        If (component = ElectricGatling) Or (.technology = "electromag") Or (.Electric) Then
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
        End If
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
       If .BurstRadius <> -1 Then
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
        If .TypeDamage2 <> "none" Then
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
       AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
       AddPCLproperty "Max Range (yards)", .MaxRange, wdText, "Disabled"
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
       AddPCLproperty "Snap Shot", .SnapShot, wdText, "Disabled"
       AddPCLproperty "# of Shots", .Shots, wdText, "Disabled"
       AddPCLproperty "Reqt. Loaders", .Loaders, wdText, "Disabled"
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
       AddPCLproperty "Cost Per Shot", "$" & Format(.CPS, "standard"), wdText, "Disabled"
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Weight Per Shot", Format(.WPS, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Volume Per Shot", Format(.VPS, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
        
    Case AntiBlastMagazine
       AddPCLproperty "Statistics", "", wdText, "Disabled"
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
    
    Case UniversalMount, CasemateMount, DoorMount, Cyberslave, FullStabilizationGear, PartialStabilizationGear
       AddPCLproperty "Statistics", "", wdText, "Disabled"
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
        
    
    
    Case WeaponBay, HardPoint
       AddPCLproperty "Settings", "", wdText, "Disabled"
       AddPCLproperty "Index", .index, wdNumber, "Index"
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
       AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
       AddPCLproperty "Maximum Load (lbs)", .loadcapacity, wdDouble, "LoadCapacity"
       AddPCLproperty "Statistics", "", wdText, "Disabled"
       AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
        If component = WeaponBay Then
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
        End If
       
    
    Case Ammunition
       AddPCLproperty "Settings", "", wdText, "Disabled"
       AddPCLproperty "# of Shots", .NumShots, wdNumber, "NumShots"
       AddPCLproperty "Lock Ammo Settings", .Locked, wdBool, "Locked"
       AddPCLproperty "Statistics", "", wdText, "Disabled"
       AddPCLproperty "Ammo Type", .Ammunitiontype, wdText, "Disabled"
       AddPCLproperty "CPS", "$" & Format(.CPS, "standard"), wdText, "Disabled"
       AddPCLproperty "WPS", Format(.WPS, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "VPS", Format(.VPS, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

        
End Select

End With
End Sub

Private Sub ShowPropsForBody()
    With m_oCurrentVeh.Components(BODY_KEY)
        AddPCLproperty "Settings", "", wdText, "Disabled"
        AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
        AddPCLproperty "Compartmentalization", .Compartmentalization, wdList, "Compartmentalization", "none", "heavy", "total"
        AddPCLproperty "Flexibody Option", .FlexibodyOption, wdBool, "FlexibodyOption"
        AddPCLproperty "Improved Flexibody Suspension", .ImprovedSuspension, wdBool, "ImprovedSuspension"
        AddPCLproperty "Lifting Body", .liftingbody, wdBool, "LiftingBody"
        AddPCLproperty "Top Deck", .TopDeck, wdBool, "Topdeck"
        AddPCLproperty "% Covered Deck", .PercentCovered, wdNumber, "PercentCovered"
        AddPCLproperty "% Flight Deck", .PercentFlightDeck, wdNumber, "PercentFlightDeck"
        AddPCLproperty "Flight Deck Option", .flightdeckoption, wdList, "flightdeckoption", "none", "landing pad", "angled flight deck"
        AddPCLproperty "Slope Right", .SlopeR, wdList, "sloper", "none", "30 degrees", "60 degrees"
        AddPCLproperty "Slope Left", .slopel, wdList, "slopel", "none", "30 degrees", "60 degrees"
        AddPCLproperty "Slope Front", .slopef, wdList, "slopeF", "none", "30 degrees", "60 degrees"
        AddPCLproperty "Slope Back", .slopeb, wdList, "slopeb", "none", "30 degrees", "60 degrees"
        AddPCLproperty "Empty Space", .EmptySpace, wdDouble, "EmptySpace"
        AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
        AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
        AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
        AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
        AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
        AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
        
        AddPCLproperty "Statistics", "", wdText, "Disabled"
        AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
        AddPCLproperty "Top Deck Area", Format(.TotalDeckArea, "standard") & " sq ft", wdText, "Disabled"
        AddPCLproperty "Flight Deck Length", Format(.flightdecklength, "standard") & " ft", wdText, "Disabled"
        AddPCLproperty "Flight Deck Area", Format(.FlightDeckArea, "standard") & " sq ft", wdText, "Disabled"
        AddPCLproperty "Covered Deck Area", Format(.covereddeckarea, "standard") & " sq ft", wdText, "Disabled"
        AddPCLproperty "Deck Cost", "$" & Format(.DeckCost, "standard"), wdText, "Disabled"
        AddPCLproperty "Deck Weight", Format(.DeckWeight, "standard") & " lbs", wdText, "Disabled"
        AddPCLproperty "Body Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
        AddPCLproperty "Body Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
        AddPCLproperty "Body Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
        AddPCLproperty "Body Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
        AddPCLproperty "Access Space", Format(.AccessSpace, "standard") & " cf", wdText, "Disabled"
        AddPCLproperty "Minimum Volume", .MinimumVolume, wdText, "Disabled"
        AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
    End With
    
End Sub
Private Sub ShowPropsForSubAssemblies(ByVal component As Integer, ByVal Key As String)

With m_oCurrentVeh.Components(Key)

' Fill the window with properties for the correct Collection item
Select Case component
        
    
        Case Wheel
            AddPCLproperty "Settings", "", wdText, "Disabled"
            AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
            AddPCLproperty "Wheel Type", .subtype, wdList, "Subtype", "standard", "small", "heavy", "railway", "off-road", "retractable"
            AddPCLproperty "# of Wheels", .Quantity, wdNumber, "Quantity"
            AddPCLproperty "Improved Suspension", .ImprovedSuspension, wdBool, "ImprovedSuspension"
            AddPCLproperty "Retract Location", .RetractLocation, wdList, "RetractLocation", "none", "body", "body & wings"
            AddPCLproperty "Wheel Blades", .Wheelblades, wdList, "Wheelblades", "none", "fixed", "rectractable"
            AddPCLproperty "Snow Tires", .snowtires, wdBool, "Snowtires"
            AddPCLproperty "Racing Tires", .racingtires, wdBool, "RacingTires"
            AddPCLproperty "Puncture Resistant", .PunctureResistant, wdBool, "PunctureResistant"
            AddPCLproperty "Improved Brakes", .ImprovedBrakes, wdBool, "ImprovedBrakes"
            AddPCLproperty "All Wheel Steering", .AllwheelSteering, wdBool, "AllWheelSteering"
            AddPCLproperty "Smart Wheels", .Smartwheels, wdBool, "SmartWheels"
            AddPCLproperty "Empty Space", .EmptySpace, wdDouble, "EmptySpace"
            'note, no empty space allowed
            AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
            AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
            AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
            AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
            AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
            AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
            AddPCLproperty "Statistics", "", wdText, "Disabled"
            AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
            AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
            AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
            AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
            AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
            AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
    Case Skid
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "# of Skids", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Improved Suspension", .ImprovedSuspension, wdBool, "ImprovedSuspension"
           AddPCLproperty "Retract Location", .RetractLocation, wdList, "RetractLocation", "none", "body", "body & wings"
            'note, no empty space allowed
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
                
        Case Track
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Track Type", .subtype, wdList, "SubType", "tracks", "halftracks", "skitracks"
           AddPCLproperty "# of Tracks", .Quantity, wdList, "Quantity", 2, 4
           AddPCLproperty "Improved Suspension", .ImprovedSuspension, wdBool, "ImprovedSuspension"
            'note, no empty space allowed
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
                                   
        Case Arm
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Index", .index, wdNumber, "Index"
           AddPCLproperty "Orientation", .Orientation, wdList, "Orientation", "top", "underside", "front", "back", "left", "right"
           AddPCLproperty "Empty Space", .EmptySpace, wdDouble, "EmptySpace"
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Access Space", Format(.AccessSpace, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
            
        Case Leg
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Index", .index, wdNumber, "Index"
           AddPCLproperty "Improved Suspension", .ImprovedSuspension, wdBool, "ImprovedSuspension"
           AddPCLproperty "Empty Space", .EmptySpace, wdDouble, "EmptySpace"
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Access Space", Format(.AccessSpace, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
                                              
        Case Wing
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Index", .index, wdNumber, "Index"
           AddPCLproperty "Orientation", .Orientation, wdList, "Orientation", "left", "right"
           AddPCLproperty "Wing Type", .subtype, wdList, "SubType", "standard", "STOL", "biplane", "triplane", "high agility", "flarecraft", "stub"
           AddPCLproperty "Controlled Instability", .ControlledInstability, wdBool, "ControlledInstability"
           AddPCLproperty "Folding Wings", .Folding, wdBool, "Folding"
           AddPCLproperty "Variable Sweep Wings", .VariableSweep, wdList, "VariableSweep", "none", "manual", "automatic"
           AddPCLproperty "Empty Space", .EmptySpace, wdDouble, "EmptySpace"
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Access Space", Format(.AccessSpace, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
            
        Case AutogyroRotor, TTRotor, CARotor, MMRotor
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Index", .index, wdNumber, "Index"
           AddPCLproperty "Controlled Instability", .ControlledInstability, wdBool, "ControlledInstability"
           AddPCLproperty "Folding Rotors", .Folding, wdBool, "Folding"
            'note, no empty space allowed
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
            
        Case Hydrofoil
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Index", .index, wdNumber, "Index"
            'note, no empty space allowed
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
            'access space because Aquatic propulsion can be placed in them
           AddPCLproperty "Access Space", Format(.AccessSpace, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
            
        Case Hovercraft
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Index", .index, wdNumber, "Index"
           AddPCLproperty "Hovercraft Type", .subtype, wdList, "SubType", "GEV skirt", "SEV sidewalls"
            'note, no empty space allowed
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
            
        Case Superstructure
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Index", .index, wdNumber, "Index"
           AddPCLproperty "Orientation", .Orientation, wdList, "Orientation", "top", "underside", "front", "back", "left", "right"
           AddPCLproperty "Compartmentalization", .Compartmentalization, wdList, "Compartmentalization", "none", "heavy", "total"
           AddPCLproperty "Top Deck", .TopDeck, wdBool, "TopDeck"
           AddPCLproperty "% Covered Deck", .PercentCovered, wdNumber, "PercentCovered"
           AddPCLproperty "% Flight Deck", .PercentFlightDeck, wdNumber, "PercentFlightDeck"
           AddPCLproperty "Flight Deck Option", .flightdeckoption, wdList, "FlightDeckOption", "none", "landing pad", "angled flight deck"
           AddPCLproperty "Slope Right", .SlopeR, wdList, "sloper", "none", "30 degrees", "60 degrees"
           AddPCLproperty "Slope Left", .slopel, wdList, "slopel", "none", "30 degrees", "60 degrees"
           AddPCLproperty "Slope Front", .slopef, wdList, "slopeF", "none", "30 degrees", "60 degrees"
           AddPCLproperty "Slope Back", .slopeb, wdList, "slopeb", "none", "30 degrees", "60 degrees"
           AddPCLproperty "Empty Space", .EmptySpace, wdDouble, "EmptySpace"
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
           AddPCLproperty "Top Deck Area", Format(.TotalDeckArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Flight Deck Length", Format(.flightdecklength, "standard") & " ft", wdText, "Disabled"
           AddPCLproperty "Flight Deck Area", Format(.FlightDeckArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Covered Deck Area", Format(.covereddeckarea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Deck Cost", "$" & Format(.DeckCost, "standard"), wdText, "Disabled"
           AddPCLproperty "Deck Weight", Format(.DeckWeight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Access Space", Format(.AccessSpace, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
            
        Case OpenMount
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Index", .index, wdNumber, "Index"
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Orientation", .Orientation, wdList, "Orientation", "top", "underside", "front", "back", "left", "right"
           AddPCLproperty "Rotation Type", .Rotation, wdList, "Rotation", "none", "full", "limited"
           AddPCLproperty "Empty Space", .EmptySpace, wdDouble, "EmptySpace"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
            
        Case Mast
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Index", .index, wdNumber, "Index"
           AddPCLproperty "# of Masts", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Height", .Height, wdNumber, "Height"
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "wood", "metal"
            'note no empty space allowed
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            

        Case Pod
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Index", .index, wdNumber, "Index"
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Orientation", .Orientation, wdList, "Orientation", "top", "underside", "front", "back", "left", "right"
           AddPCLproperty "Empty Space", .EmptySpace, wdDouble, "EmptySpace"
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Access Space", Format(.AccessSpace, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
            
        Case Turret, Popturret
           AddPCLproperty "Settings", "", wdText, "Disabled"
            'note: only turrets and popturrets will have a "rotation space" statistic
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Index", .index, wdNumber, "Index"
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Orientation", .Orientation, wdList, "Orientation", "top", "underside", "front", "back", "left", "right"
           AddPCLproperty "Rotation", .Rotation, wdList, "Rotation", "none", "limited", "full"
           AddPCLproperty "Compartmentalization", .Compartmentalization, wdList, "Compartmentalization", "none", "heavy", "total"
           AddPCLproperty "Slope Right", .SlopeR, wdList, "sloper", "none", "30 degrees", "60 degrees"
           AddPCLproperty "Slope Left", .slopel, wdList, "slopel", "none", "30 degrees", "60 degrees"
           AddPCLproperty "Slope Front", .slopef, wdList, "slopeF", "none", "30 degrees", "60 degrees"
           AddPCLproperty "Slope Back", .slopeb, wdList, "slopeb", "none", "30 degrees", "60 degrees"
           AddPCLproperty "Empty Space", .EmptySpace, wdDouble, "EmptySpace"
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
           AddPCLproperty "Rotation Space", .RotationSpace, wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Access Space", Format(.AccessSpace, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
            
        Case Gasbag
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Index", .index, wdNumber, "Index"
            'note no empty space allowed
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
            
         Case Cargo
           AddPCLproperty "Settings", "", wdText, "Disabled"
           'AddPCLproperty "Index", .Index, wdNumber, "Index"
           AddPCLproperty "Cargo Type", .subtype, wdList, "Subtype", "standard", "hidden", "open"
           AddPCLproperty "Cargo Room", .CargoSpace, wdDouble, "CargoSpace"
           AddPCLproperty "Empty Weight", .Weight, wdDouble, "Weight"
           AddPCLproperty "Weight Per cf", .WeightPerCubicFoot, wdDouble, "WeightPerCubicFoot"
            'note no empty space allowed
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           'AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
           AddPCLproperty "Compartment Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Cargo Weight", Format(.CargoWeight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
            
            
        Case equipmentPod
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Index", .index, wdNumber, "Index"
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Empty Space", .EmptySpace, wdDouble, "EmptySpace"
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Access Space", Format(.AccessSpace, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
       Case SideCar
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Index", .index, wdNumber, "Index"
            'note, no empty space allowed
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Access Space", Format(.AccessSpace, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
       
       Case SolarPanel
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           'AddPCLproperty "Index", .Index, wdNumber, "Index"
           AddPCLproperty "Surface Area", .SurfaceArea, wdDouble, "SurfaceArea"
           AddPCLproperty "Retractable?", .Retractable, wdBool, "Retractable"
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           'AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
      Case SolarCellArray
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Percent Area Covered", .PercentCovered, wdNumber, "PercentCovered"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kw", wdText, "Disabled"
           'AddPCLproperty "Endurance", .Endurance & " yrs", wdText, "Disabled"
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
    End Select
End With
End Sub
         
Private Sub ShowPropsForPropulsion(ByVal component As Integer, ByVal Key As String)
'////////////////////////////////////////////
'Propulsion Systems
'////////////////////////////////////////////
With m_oCurrentVeh.Components(Key)

' Fill the window with properties for the correct Collection item
Select Case component
        Case WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, TrackedDrivetrain, LegDrivetrain, FlexibodyDrivetrain
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           If component = LegDrivetrain Then
                AddPCLproperty "Motive Power (per motor)", .motivepower, wdDouble, "MotivePower"
           Else
                AddPCLproperty "Motive Power", .motivepower, wdDouble, "MotivePower"
           End If
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
            If component = LegDrivetrain Then
               AddPCLproperty "Volume per leg:", Format(.Volume, "standard"), wdText, "Disabled"
            Else
               AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
            End If
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
            
        Case OrnithopterDrivetrain, TTRRotorDrivetrain, CARRotorDrivetrain
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           If component = OrnithopterDrivetrain Then
                AddPCLproperty "Motive Power (per motor)", .motivepower, wdDouble, "MotivePower"
           Else
                AddPCLproperty "Motive Power", .motivepower, wdDouble, "MotivePower"
           End If
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Lift", Format(.Lift, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            

        Case MMRRotorDrivetrain
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Tilt Rotor?", .TiltRotor, wdBool, "TiltRotor"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "Motive Power", .motivepower, wdDouble, "MotivePower"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Lift", Format(.Lift, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
            
        Case AerialPropeller
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "Motive Power", .motivepower, wdDouble, "MotivePower"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           'AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           'AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           'AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
            
        Case DuctedFan
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Motive Power", .motivepower, wdDouble, "MotivePower"
           AddPCLproperty "Hover Fan", .HoverFan, wdBool, "HoverFan"
           AddPCLproperty "Lift Engine", .LiftEngine, wdBool, "LiftEngine"
           AddPCLproperty "Vectored Thrust", .VectoredThrust, wdBool, "VectoredThrust"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
            
            
        Case PaddleWheel, ScrewPropeller, lightScrewPropeller, DuctedPropeller, Hydrojet, MHDTunnel
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "Motive Power", .motivepower, wdDouble, "MotivePower"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Aquatic Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
            
        Case RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Animal Type", .subtype, wdList, "SubType", "Land Animal", "Swimming Animal", "Flying Animal"
           AddPCLproperty "Animal Description", .AnimalDescription, wdText, "AnimalDescription"
           AddPCLproperty "Strength per Animal", .BeastST, wdNumber, "BeastST"
           AddPCLproperty "Hexes Per Animal", .Hexes, wdNumber, "Hexes"
           AddPCLproperty "Number of Animals", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
            If .subtype = "Land Animal" Then
               AddPCLproperty "Motive Power", .motivepower, wdText, "Disabled"
            Else
               AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
            End If
           AddPCLproperty "Total Hexes of Animals", .TotalHexes, wdText, "Disabled"
           AddPCLproperty "Move per Animal", .Move, wdText, "Disabled"
           AddPCLproperty "Speed per Animal", .Speed, wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
        

        
        Case RowingPositions
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "# of Positions", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Avg. ST per Position", .RowerST, wdNumber, "RowerST"
           AddPCLproperty "DR per Position", .dr, wdNumber, "DR"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            

        Case FullRig, SquareRig, ForeandAftRig, AerialSail, AerialSailForeAftRig
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Sail Material", .material, wdList, "Material", "cloth", "synthetic", "bioplas"
           AddPCLproperty "Wind", .Wind, wdList, "Wind", "calm", "light air", "light breeze", "gentle breeze", "moderate breeze", "fresh breeze", "strong breeze", "gale force winds"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
                       
        Case lightSail
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "Sail Size (sq mi)", .SurfaceArea, wdDouble, "SurfaceArea"
           AddPCLproperty "AU Distance", .AUDistance, wdDouble, "AUDistance"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Motive Thrust", Format(.Thrust, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
        Case Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Thrust (lbs)", .DesiredThrust, wdDouble, "DesiredThrust"
           AddPCLproperty "Afterburner", .Afterburner, wdBool, "Afterburner"
           If component <> Ramjet Then '//ramjets cant be lift engines because they need air travelling through them in forward motion
                AddPCLproperty "Lift Engine", .LiftEngine, wdBool, "LiftEngine"
           End If
           AddPCLproperty "Vectored Thrust", .VectoredThrust, wdBool, "VectoredThrust"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Fuel Consumption", Format(.FuelConsumption, "standard") & " gph", wdText, "Disabled"
           AddPCLproperty "AB Thrust", Format(.ABThrust, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "AB Lit Fuel Consumption", Format(.ABConsumption, "standard") & " gph", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
            
        Case FusionAirRam 'only jet engine that cant use afterburner
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Thrust (lbs)", .DesiredThrust, wdDouble, "DesiredThrust"
           AddPCLproperty "Lift Engine", .LiftEngine, wdBool, "LiftEngine"
           AddPCLproperty "Vectored Thrust", .VectoredThrust, wdBool, "VectoredThrust"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
           
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Endurance", Format(.FuelConsumption, "standard") & " yrs", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
            
        Case StandardThruster, SuperThruster, MegaThruster
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Thrust (lbs)", .DesiredThrust, wdDouble, "DesiredThrust"
           AddPCLproperty "Lift Engine", .LiftEngine, wdBool, "LiftEngine"
           AddPCLproperty "Vectored Thrust", .VectoredThrust, wdBool, "VectoredThrust"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Power", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
           
        Case LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Thrust (lbs)", .DesiredThrust, wdDouble, "DesiredThrust"
           AddPCLproperty "Lift Engine", .LiftEngine, wdBool, "LiftEngine"
           AddPCLproperty "Vectored Thrust", .VectoredThrust, wdBool, "VectoredThrust"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Fuel Consumption", Format(.FuelConsumption, "standard") & " gph", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
        Case AntimatterPion
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Thrust (lbs)", .DesiredThrust, wdDouble, "DesiredThrust"
           AddPCLproperty "Lift Engine", .LiftEngine, wdBool, "LiftEngine"
           AddPCLproperty "Vectored Thrust", .VectoredThrust, wdBool, "VectoredThrust"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Antimatter Fuel Consumption", Format(.FuelConsumption, "standard") & " grams per hour", wdText, "Disabled"
           AddPCLproperty "Hydrogen Fuel Consumption", Format(.FuelConsumption2, "standard") & " gph", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
        Case SolidRocketEngine
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Thrust (lbs)", .DesiredThrust, wdDouble, "DesiredThrust"
           AddPCLproperty "Burn Time (mins)", .BurnTime, wdDouble, "BurnTime"
           AddPCLproperty "Lift Engine", .LiftEngine, wdBool, "LiftEngine"
           AddPCLproperty "Vectored Thrust", .VectoredThrust, wdBool, "VectoredThrust"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
            
        Case OrionEngine
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Pulse Rate (bps)", .PulseRate, wdDouble, "PulseRate"
           AddPCLproperty "Bomb Size (kt)", .BombSize, wdDouble, "BombSize"
           AddPCLproperty "# of Bombs", .NumBombs, wdNumber, "NumBombs"
           AddPCLproperty "Lift Engine", .LiftEngine, wdBool, "LiftEngine"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Thrust Time (secs)", .ThrustTime, wdText, "Disabled"
           AddPCLproperty "Bomb Weight", .BombWeight, wdText, "Disabled"
           AddPCLproperty "Bomb Cost", .BombCost, wdText, "Disabled"
           AddPCLproperty "Bomb Volume", .BombVolume, wdText, "Disabled"
           AddPCLproperty "Total Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Total Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Total Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
            
        Case MagLevLifter
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Lift (lbs)", .DesiredLift, wdDouble, "DesiredLift"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Total lift", Format(.Lift, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Power", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
        Case JumpDrive, TeleportationDrive
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Capacity (tons)", .DesiredCapacity, wdDouble, "DesiredCapacity"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Total Capacity", Format(.capacity, "standard") & " tons", wdText, "Disabled"
           AddPCLproperty "Power", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
            
        Case Hyperdrive
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Capacity (tons)", .DesiredCapacity, wdDouble, "DesiredCapacity"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Total Capacity", Format(.capacity, "standard") & " tons", wdText, "Disabled"
           AddPCLproperty "Initial Power", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Sustained Power", Format(.SustainedPower, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
        Case WarpDrive
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Warp Thrust Factor", .DesiredCapacity, wdDouble, "DesiredCapacity"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Total WTF", Format(.capacity, "standard") & " WTF", wdText, "Disabled"
           AddPCLproperty "Power", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
    
        Case SubQuantumConveyor, QuantumConveyor, TwoQuantumConveyor
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Max Transport Weight (lbs)", .DesiredCapacity, wdDouble, "desiredCapacity"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Total Transport Weight", Format(.capacity, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Power", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
    
    '/////////////////////////////////////////
    ' Aerostatic Lift Systems
        Case HotAir, Hydrogen, Helium
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Useful Static Lift (lbs)", .Lift, wdDouble, "Lift"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
            
        
        Case ContraGravGenerator
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Lift", .DesiredLift, wdDouble, "DesiredLift"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Total Lift", Format(.Lift, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Power", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
End Select
End With
End Sub

Private Sub ShowPropsforInstruments(ByVal component As Integer, ByVal Key As String)
'///////////////////////////////////////////
'Instruments and Electronics
With m_oCurrentVeh.Components(Key)

Select Case component
' Fill the window with properties for the correct Collection item
    Case RadioDirectionFinder, RadioCommunicator, TightBeamRadio, VLFRadio, CellularPhone, CellularPhonewithRadio, RadioJammer, ElfReceiver, LaserCommunicator, NeutrinoCommunicator, GravityRippleCommunicator
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Desired Range", .DesiredRange, wdList, "DesiredRange", "short", "medium", "long", "very long", "extreme"
           AddPCLproperty "Sensitivity", .Sensitivity, wdList, "Sensitivity", "normal", "sensitive", "very sensitive"
           AddPCLproperty "FTL", .FTL, wdBool, "FTL"
           AddPCLproperty "Receive Only", .ReceiveOnly, wdBool, "ReceiveOnly"
           AddPCLproperty "Scrambler", .Scrambler, wdBool, "Scrambler"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
            If .FTL = False Then
               AddPCLproperty "Actual Range", Format(.Range, "standard") & " miles", wdText, "Disabled"
            Else
               AddPCLproperty "Actual Range", Format(.Range, "standard") & " parsecs", wdText, "Disabled"
            End If
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
    
    Case Headlight, Searchlight, InfraredSearchlight
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
           If component = Headlight Then
               AddPCLproperty "Range (yards)", .Range, wdDouble, "Range"
           Else
              AddPCLproperty "Range (miles)", .Range, wdDouble, "Range"
           End If
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
    Case AstronomicalInstruments, Telescope, lightAmplification, LowlightTV, ExtendableSensorPeriscope
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
            If component = lightAmplification Then
            ElseIf component = ExtendableSensorPeriscope Then
               AddPCLproperty "Periscope Length", .Length, wdDouble, "Length"
            Else
               AddPCLproperty "Magnification", .Magnification, wdDouble, "Magnification"
            End If
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
    Case Radar, Ladar, NavigationalRadar, AntiCollisionRadar, AESA, LowResImagingRadar, HiResImagingRadar
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
           AddPCLproperty "Range", .Range, wdDouble, "Range"
            If component = NavigationalRadar Then
            ElseIf component = AntiCollisionRadar Then
            Else
               AddPCLproperty "No Targeting", .NoTargeting, wdBool, "NoTargeting"
               AddPCLproperty "Search Optimization", .SearchOption, wdList, "SearchOption", "none", "surface search", "air search"
               AddPCLproperty "FTL Option", .FTL, wdBool, "FTL"
            End If
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Scan Rating", .ScanRating, wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
    Case ActiveSonar, PassiveSonar
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
            If component = ActiveSonar Then
               AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
            End If
           AddPCLproperty "Range", .Range, wdDouble, "Range"
            If component = ActiveSonar Then
               AddPCLproperty "Active / Passive?", .ActivePassive, wdBool, "ActivePassive"
               AddPCLproperty "Depth Finding?", .DepthFinding, wdBool, "DepthFinding"
               AddPCLproperty "Dipping Sonar?", .DippingSonar, wdBool, "DippingSonar"
               AddPCLproperty "No Targeting?", .NoTargeting, wdBool, "NoTargeting"
            Else
               AddPCLproperty "Dipping Sonar?", .DippingSonar, wdBool, "DippingSonar"
               AddPCLproperty "Towed Array?", .TowedArray, wdBool, "TowedArray"
            End If
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Scan Rating", .ScanRating, wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
    Case PassiveInfrared, Thermograph, PassiveRadar, PESA
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
           AddPCLproperty "Range", .Range, wdDouble, "Range"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Scan Rating", .ScanRating, wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
    Case Geophone, MAD, MultiScanner, ChemScanner, RadScanner, BioScanner, GravScanner
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Range", .Range, wdDouble, "Range"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Scan Rating", .ScanRating, wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
    Case RangingSoundDetector, SurveillanceSoundDetector
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Sensitivity Level", .Level, wdNumber, "Level"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
    Case MeteorologicalInstruments, LowResPlanetarySurveyArray, MedResPlanetarySurveyArray, HighResPlanetarySurveyArray
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
    Case SoundSystem, FlightRecorder
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
    
    Case VehicleCamera, DigitalVehicleCamera, ReconCamera, DigitalReconCamera
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Low light?", .Lowlight, wdBool, "Lowlight"
           AddPCLproperty "Infrared?", .Infrared, wdBool, "Infrared"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
    Case NavigationInstruments, AutoPilot, IFF, Transponder, INS, GPS, MilitaryGPS, TFR
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           If component = NavigationInstruments Then
               AddPCLproperty "Precision?", .Precision, wdBool, "Precision"
           End If
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
    Case ImprovedOpticalBombSight, AdvancedOpticalBombSight, OpticalBombSight, FireDirectionCenter, HUDWAC, PupilHUDWAC, LaserRangeFinder, LaserDesignator, LaserSpotTracker
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           If component = LaserDesignator Or component = LaserRangeFinder Then
               AddPCLproperty "Range", .Range, wdDouble, "Range"
           End If
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
    Case RadarDetector, LaserSensor, LaserRadarDetector, AreaRadarJammer, _
        DeceptiveRadarJammer, InfraredJammer, ChaffDecoyDischarger, SmokeDecoyDischarger, _
        FlareDecoyDischarger, SonarDecoyDischarger, HotSmokeDecoyDischarger, _
        PrismDecoyDischarger, BlackOutGasDecoyDischarger, RadarReflector, BlipEnhancer, TEMPEST
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
            Select Case component
                Case AreaRadarJammer, DeceptiveRadarJammer, InfraredJammer
                    AddPCLproperty "Jammer Rating", .JammerRating, wdList, "JammerRating", 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
                Case RadarDetector, LaserSensor, LaserRadarDetector
                    AddPCLproperty "Advanced Version?", .ADVANCED, wdBool, "advanced"
            End Select
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
    Case DecoyChaff, DecoySmoke, DecoyFlares, DecoySonarDecoy, DecoyHotSmoke, DecoyPrism, DecoyBlackOutGas
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
         
    Case MacroFrame, MainFrame, MicroFrame, MiniComputer, SmallComputer
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Intelligence", .Intelligence, wdList, "Intelligence", "normal", "dumb", "genius"
           AddPCLproperty "Configuration", .Configuration, wdList, "Configuration", "normal", "neural-net", "sentient"
           AddPCLproperty "Compact?", .Compact, wdBool, "Compact"
           AddPCLproperty "Hardened?", .Hardened, wdBool, "Hardened"
           AddPCLproperty "High Capacity?", .HighCapacity, wdBool, "HighCapacity"
           AddPCLproperty "Dedicated?", .Dedicated, wdBool, "Dedicated"
           AddPCLproperty "Robot Brain?", .RobotBrain, wdBool, "RobotBrain"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Complexity", .complexity, wdText, "Disabled"
           If .IQ > 0 Then
                AddPCLproperty "IQ", .IQ, wdText, "Disabled"
           End If
           If .DX > 0 Then
                AddPCLproperty "DX", .DX, wdText, "Disabled"
           End If
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
                   
    Case Terminal
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
    Case DatabaseSoftware
        AddPCLproperty "Settings", "", wdText, "Disabled"
        AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
        AddPCLproperty "Size (gigs)", .gigabytes, wdDouble, "Gigabytes"
        AddPCLproperty "Statistics", "", wdText, "Disabled"
        AddPCLproperty "Complexity", .complexity, wdText, "Disabled"
        AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
    
    Case CartographySoftware, ComputerNavigationSoftware, _
        DatalinkSoftware, TransmissionProfilingSoftware, HoloventureProgram, _
        PersonalitySimulationSoftwareFull, PersonalitySimulationLimited, _
        RoutineVehicleOperationSoftwarePilot, RoutineVehicleOperationSoftwareOther
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Complexity", .complexity, wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
            
        
    Case FireDirectionSoftware, TargetingSoftware, DamageControlSoftware, _
        GunnerSoftware, RobotSkillProgramsPhysical, RobotSkillProgramsMental
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Bonus Skill", .BonusSkillPoints, wdNumber, "BonusSkillPoints"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Total Skill Points", .SkillPoints, wdText, "Disabled"
           AddPCLproperty "Complexity", .complexity, wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
     
            
    Case SurgicalInterface, InterfaceWeb, AutoInterfaceWeb, SocketInterface, NeuralInductionField
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "# of Users", .Quantity, wdNumber, "Quantity"
            If component = SocketInterface Then
               AddPCLproperty "DR", .dr, wdNumber, "DR"
               AddPCLproperty "Statistics", "", wdText, "Disabled"
               AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
               AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
               AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
            Else
               AddPCLproperty "DR", .dr, wdNumber, "DR"
               AddPCLproperty "Statistics", "", wdText, "Disabled"
               AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
               AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
               AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
               AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
               AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
               AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            End If
     
     Case DeflectorField
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "PD Bonus", "+" + Format(.PDBonus), wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           'AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           'AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           'AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
    
    Case ForceScreen, VariableForceScreen
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Screen DR", .ForceDR, wdNumber, "ForceDR"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           'AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           'AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           'AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
End Select
End With
End Sub

Private Sub ShowPropsForMiscellanous(ByVal component As Integer, ByVal Key As String)
'///////////////////////////////////////////
'Miscellanous equipment
'///////////////////////////////////////////
With m_oCurrentVeh.Components(Key)

Select Case component
' Fill the window with properties for the correct Collection item
    Case ArmMotor
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "ST", .ST, wdNumber, "ST"
           AddPCLproperty "Bad Grip?", .BadGrip, wdBool, "BadGrip"
           AddPCLproperty "Cheap?", .Cheap, wdBool, "Cheap"
           AddPCLproperty "Extendable?", .Extendable, wdBool, "Extendable"
           AddPCLproperty "Poor Coordination?", .PoorCoordination, wdBool, "PoorCoordination"
           AddPCLproperty "Striker?", .Striker, wdBool, "Striker"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
           
        
    Case FireExtinguisherSystem, FullFireSuppressionSystem, CompactFireSuppressionSystem
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
    Case BilgePump
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
    Case CompleteWorkshop, MechanicWorkshop, EngineeringWorkshop, ElectronicsWorkshop, ArmouryWorkshop, CompleteMiniWorkshop, ScienceLab, MiniMechanicWorkshop, MiniElectronicsWorkshop, MiniEngineeringWorkshop, MiniArmouryWorkshop
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
            If component = ScienceLab Then
               'AddPCLproperty "Skill", .Skill, wdList, "Skill", "astronomy", "biochemistry", "biology", "botany", "chemistry", "computer programming", "criminology", "ecology", "economics", "electronics", "engineering", "forensics", "genetics", "geology", "history", "linguistics", "literature", "mathematics", "metallurgy", "meteorology", "nuclear physics", "occultism", "physics", "physiology", "prospecting", "psychology", "research", "theology", "zoology"
               AddPCLproperty "Skill", .Skill, wdText, "Skill"
            End If
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
    Case ExtendableLadder, Crane, Winch, PowerShovel, WreckingCrane, ForkLift, VehicularBridge, LaunchCatapult, SkyHook, Bore, SuperBore, EnergyDrill, TractorBeam, PressorBeam, CombinationBeam, CraneWithElectroMagnet
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
            Select Case component
            Case ExtendableLadder, Crane, CraneWithElectroMagnet, WreckingCrane
               AddPCLproperty "Crane Height (ft)", .Height, wdNumber, "Height"
            Case PowerShovel, Winch, ForkLift, TractorBeam, PressorBeam, CombinationBeam
               AddPCLproperty "ST", .ST, wdNumber, "ST"
            Case VehicularBridge
               AddPCLproperty "Length (yds)", .Length, wdDouble, "Length"
               AddPCLproperty "Max Supported Weight", .DesiredWeight, wdDouble, "DesiredWeight"
            Case Bore, SuperBore
               AddPCLproperty "Tunneling Ability Per Hour (cf)", .TunnelingAbility, wdDouble, "TunnelingAbility"
            Case SkyHook, LaunchCatapult
            End Select
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
    Case OperatingRoom, StretcherPallet, EmergencySupportUnit, EmergencylightsandSiren, CryonicCapsule, Automed, DiagnosisTable
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
            If component = OperatingRoom Then
               AddPCLproperty "# of Operating Tables", .OperatingTables, wdNumber, "OperatingTables"
            End If
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
            
    Case Stage, Hall, BarRoom, ConferenceRoom, MovieScreenandProjector, MovieScreenandProjectorSmall, HoloventureZone
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
            Select Case component
                Case Stage, Hall, BarRoom, ConferenceRoom, HoloventureZone
                   AddPCLproperty "Floor Area", .FloorArea, wdDouble, "FloorArea"
                Case Else
                    AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
            End Select
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        'Note: door and hatch have been remove below.  User should enter these in as "Details" in the options dialog
    Case CargoRamp, Airlock, MembraneAirlock, Forcelock, PassageTube, ArmoredPassageTube
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
            Select Case component
                Case Airlock, MembraneAirlock
                   AddPCLproperty "# People Supported", .Rating, wdNumber, "Rating"
                Case Else
            End Select
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
    Case TeleportProjector
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "# Hexes", .HexCapacity, wdNumber, "HexCapacity"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
               
    Case BrigsandRestraints, BurglarAlarm, HighSecurityAlarm, MutableLicensePlate, OilSprayer, PaintSprayer, SmokeScreen, SpikeDropper
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
    Case VehicleBay, HangerBay, DryDock, SpaceDock, ExternalCradle
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
            If component = ExternalCradle Then
               AddPCLproperty "Total Craft Weight", .CraftWeight, wdDouble, "CraftWeight"
            Else
               AddPCLproperty "Total Craft Weight", .CraftWeight, wdDouble, "CraftWeight"
               AddPCLproperty "Cubic Feet of Craft", .CubicFeetCraft, wdDouble, "CubicFeetCraft"
            End If
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
    Case ArrestorHook, VehicularParachute
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
            If component = VehicularParachute Then
               AddPCLproperty "Rated Weight", .RatedWeight, wdDouble, "RatedWeight"
            End If
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
            
    Case RefuellingProbe, RefuellingDrogue, FuelElectrolysisSystem, HydrogenFuelScoop, AtmosphereProcessor
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
            If (component = FuelElectrolysisSystem) Or (component = AtmosphereProcessor) Then
               AddPCLproperty "Processing Capacity (gallons)", .capacity, wdDouble, "Capacity"
            End If
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
    Case NuclearDamper
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Field Radius (mi)", .Radius, wdDouble, "Radius"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
    Case SmallRealityStabilizer, MediumRealityStabilizer, HeavyRealityStabilizer
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
    Case ModularSocket
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Rated Volume", .RatedVolume, wdDouble, "RatedVolume"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
            
            
    Case Module
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Waste Weight", .WasteWeight, wdDouble, "WasteWeight"
           AddPCLproperty "Waste Volume", .WasteVolume, wdDouble, "WasteVolume"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
            
            
            
End Select
End With
End Sub
        
Private Sub ShowPropsForPowerandFuel(ByVal component As Integer, ByVal Key As String)
'///////////////////////////////////////////
'Power and Fuel
'//////////////////////////////////////////


Dim listarray() As String
ReDim listarray(1)

With m_oCurrentVeh.Components(Key)

' Fill the window with properties for the correct Collection item

Select Case component

 Case MuscleEngine
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Maximum Output", .MaxOutput, wdDouble, "MaxOutPut"
           AddPCLproperty "Combined Operator ST", .CombinedST, wdNumber, "CombinedST"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
            
Case EarlySteamEngine, ForcedDraftSteamEngine, TripleExpansionSteamEngine, SteamTurbine
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Output", .DesiredOutput, wdDouble, "DesiredOutput"
           If component = SteamTurbine Then
                AddPCLproperty "Fuel Type", .Fueltype, wdList, "FuelType", "coal", "diesel fuel"
            Else
                AddPCLproperty "Fuel Type", .Fueltype, wdList, "FuelType", "coal", "wood"
            End If
           
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Fuel Consumption", .FuelConsumption, wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
            
 Case GasolineEngine, HPGasolineEngine, TurboGasolineEngine, SuperGasolineEngine, TurboHPGasolineEngine, _
    SuperHPGasolineEngine, StandardDieselEngine, TurboStandardDieselEngine, MarineDieselEngine, _
    HPDieselEngine, TurboHPDieselEngine, CeramicEngine, TurboCeramicEngine, SuperCeramicEngine, _
    HPCeramicEngine, TurboHPCeramicEngine, SuperHPCeramicEngine, HydrogenCombustionEngine
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Output", .UnModifiedOutput, wdDouble, "UnModifiedOutput"
           'hydrogen combustion
            If component = HydrogenCombustionEngine Then
                AddPCLproperty "Fuel Type", .Fueltype, wdList, "FuelType", "hydrogen"
            'aviation fuels
            ElseIf (component = HPCeramicEngine) Or (component = TurboHPCeramicEngine) Or _
                (component = SuperHPCeramicEngine) Or (component = HPGasolineEngine) Or _
                (component = TurboHPGasolineEngine) Or (component = SuperHPGasolineEngine) Then
                AddPCLproperty "Fuel Type", .Fueltype, wdList, "FuelType", "aviation gas"
            'multifuels
            ElseIf (component = CeramicEngine) Or (component = TurboCeramicEngine) Or (component = SuperCeramicEngine) Then
                AddPCLproperty "Fuel Type", .Fueltype, wdList, "FuelType", "gasoline", "diesel fuel", "aviation gas", "ethanol", "methanol"
            'diesels with alcohol / propane potential
            ElseIf (component = TurboStandardDieselEngine) Or (component = TurboHPDieselEngine) Or _
                (component = MarineDieselEngine) Or (component = StandardDieselEngine) Or _
                (component = HPDieselEngine) Then
                AddPCLproperty "Fuel Type", .Fueltype, wdList, "FuelType", "diesel fuel", "propane", "ethanol", "methanol"
            'gasolines with alcohol / propane potential
            ElseIf (component = GasolineEngine) Or (component = TurboGasolineEngine) Or _
            (component = SuperGasolineEngine) Then
                AddPCLproperty "Fuel Type", .Fueltype, wdList, "FuelType", "gasoline", "propane", "ethanol", "methanol"
            End If
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Fuel Consumption", .FuelConsumption, wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
    

Case FuelCell, HPGasTurbine, StandardMHDTurbine, HPMHDTurbine
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Output", .DesiredOutput, wdDouble, "DesiredOutPut"
           AddPCLproperty "Closed Cycle?", .ClosedCycle, wdBool, "ClosedCycle"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Fuel Consumption", .FuelConsumption, wdText, "Disabled"
           AddPCLproperty "LOX Consumption", .LOXConsumption, wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
Case StandardGasTurbine, OptimizedGasTurbine
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Output", .DesiredOutput, wdDouble, "DesiredOutPut"
           AddPCLproperty "Fuel Type", .Fueltype, wdList, "FuelType", "gasoline", "diesel fuel", "alcohol", "aviation gas", "jet fuel"
           AddPCLproperty "Closed Cycle?", .ClosedCycle, wdBool, "ClosedCycle"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Fuel Consumption", .FuelConsumption, wdText, "Disabled"
           AddPCLproperty "LOX Consumption", .LOXConsumption, wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
    
    'only difference between FissionReactor and the others is the Uranium Fuel Rods
Case FissionReactor
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Output", .DesiredOutput, wdDouble, "DesiredOutput"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Endurance", .Endurance & " yrs", wdText, "Disabled"
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Fuel Rods Installed", .FuelConsumption, wdText, "Disabled"
           AddPCLproperty "Fuel Rod Added Cost", .FuelCost, wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
    

Case RTGReactor, NPU, FusionReactor, AntimatterReactor, TotalConversionPowerPlant, CosmicPowerPlant
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Output", .DesiredOutput, wdDouble, "DesiredOutPut"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Endurance", .Endurance & " yrs", wdText, "Disabled"
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
    
    
Case Soulburner, ElementalFurnace, ManaEngine
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Output", .DesiredOutput, wdDouble, "DesiredOutPut"
           AddPCLproperty "Cost for Magic", .MagicCost, wdDouble, "MagicCost"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
    
Case Carnivore, Herbivore, Omnivore, Vampire
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Output", .DesiredOutput, wdDouble, "DesiredOutPut"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
    
Case ClockWork
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Stored Capacity (kWs)", .DesiredOutput, wdDouble, "DesiredOutPut"
           AddPCLproperty "Powered Rewind Mechanism?", .PoweredRewinder, wdBool, "PoweredRewinder"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Rewind Motor ST", .MotorST, wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            

Case LeadAcidBattery, AdvancedBattery, Flywheel, RechargeablePowerCell, PowerCell
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Stored Capacity (kWs)", .DesiredOutput, wdDouble, "DesiredOutPut"
           If (component = PowerCell) Or (component = RechargeablePowerCell) Then
                AddPCLproperty "Cell Type", .CellType, wdList, "CellType", "custom", "AA", "A", "B", "C", "D", "E"
           End If
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
    
Case AntiMatterBay
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Capacity (grams)", .capacity, wdDouble, "Capacity"
           AddPCLproperty "Failsafe Points", .FailSafePoints, wdNumber, "FailSafePoints"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Fuel Cost", .FuelCost, wdText, "Disabled"
           AddPCLproperty "Fuel Weight", .FuelWeight, wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            

Case StandardTank, lightTank, UltralightTank, StandardSelfSealingTank, lightSelfSealingTank, UltralightSelfSealingTank
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Capacity (gallons)", .capacity, wdDouble, "Capacity"
           AddPCLproperty "Fuel Type", .Fuel, wdList, "Fuel", "ethanol", "methanol", "aviation gas", "cadmium", "diesel", "gasoline", "jet fuel", "rocket fuel", "water", "hydrogen", "metal/LOX", "oxygen (LOX)", "propane/LNG"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Fire", .Fire, wdText, "Disabled"
           AddPCLproperty "Tank Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Tank Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Tank Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Fuel Fire", .FuelFire, wdText, "Disabled"
           AddPCLproperty "Fuel Cost", .FuelCost, wdText, "Disabled"
           AddPCLproperty "Fuel Weight", .FuelWeight, wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
    
Case CoalBunker, WoodBunker
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Capacity (cubic ft.)", .capacity, wdDouble, "Capacity"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Fuel Cost", .FuelCost, wdText, "Disabled"
           AddPCLproperty "Fuel Weight", .FuelWeight, wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
    
'Case Water, Wood, Coal, Gasoline, Diesel, AviationGas, JetFuel, Propane, LiquifiedNaturalGas, EthanolAlchohol, MethanolAlchohol, LiquidHydrogen, LiquidOxygen, Cadmium, MetalLOX, RocketFuel, AntiMatter
    
Case ElectricContactPower
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Power Drawn", .DesiredOutput, wdDouble, "DesiredOutPut"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
    
Case LaserBeamedPowerReceiver, MaserBeamedPowerReceiver
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Max Power", .DesiredOutput, wdDouble, "DesiredOutPut"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
           
Case NitrousOxideBooster
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Max Boost Length", .MaxBoostLength & " seconds", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
           
Case Snorkel

           listarray = .FillCombustionEngineList
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Assigned Power Plants", .PowerPlants, wdList, "PowerPlants", listarray
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

End Select

End With
End Sub
   
Private Sub ShowPropsForArmor(ByVal component As Integer, ByVal ComponentsParent As Integer, ByVal Key As String)
Dim listarray() As String

With m_oCurrentVeh.Components(Key)

' Fill the window with properties for the correct Collection item

    '///////////////////////////////////////////
    'Armor
     Select Case component
        
        Case ArmorBasicFacing
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           listarray = .FillMaterial
           AddPCLproperty "Material", .material, wdList, "Material", listarray
           listarray = .FillQuality(.material)
           AddPCLproperty "Quality", .Quality, wdList, "Quality", listarray
           AddPCLproperty "Surface Coating", .coating, wdList, "Coating", "none", "reflective", "retro-reflective"
           AddPCLproperty "Radiation Shielding", .radiation, wdBool, "Radiation"
           AddPCLproperty "Thermal Superconductor", .thermal, wdBool, "Thermal"
           AddPCLproperty "Reactive Armor Plating", .rap, wdBool, "RAP"
           AddPCLproperty "Electrified", .electrified, wdBool, "Electrified"
           AddPCLproperty "DR (Right)", .dr1, wdNumber, "DR1"
           AddPCLproperty "DR (Left)", .dr2, wdNumber, "DR2"
           AddPCLproperty "DR (Front)", .dr3, wdNumber, "DR3"
           AddPCLproperty "DR (Back)", .dr4, wdNumber, "DR4"
           AddPCLproperty "DR (Top)", .dr5, wdNumber, "DR5"
           If ComponentsParent = Body Then
               AddPCLproperty "DR (Bottom)", .dr6, wdNumber, "DR6"
           End If
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Average DR", .AverageDR, wdText, "Disabled"  'MAKE SURE IM Calcing this properly IN THE CLASS depending on whether im dealling with 5 sides or 6!
           AddPCLproperty "Effective DR (Right)", .EffectiveDR1, wdText, "Disabled"
           AddPCLproperty "Effective DR (Left)", .EffectiveDR2, wdText, "Disabled"
           AddPCLproperty "Effective DR (Front)", .EffectiveDR3, wdText, "Disabled"
           AddPCLproperty "Effective DR (Back)", .EffectiveDR4, wdText, "Disabled"
           AddPCLproperty "Effective DR (Top)", .EffectiveDR5, wdText, "Disabled"
           If ComponentsParent = Body Then
               AddPCLproperty "Effective DR (Bottom)", .EffectiveDR6, wdText, "Disabled"
           End If
           AddPCLproperty "PD (Right)", .PD1, wdText, "Disabled"
           AddPCLproperty "PD (Left)", .PD2, wdText, "Disabled"
           AddPCLproperty "PD (Front)", .PD3, wdText, "Disabled"
           AddPCLproperty "PD (Back)", .PD4, wdText, "Disabled"
           AddPCLproperty "PD (Top)", .PD5, wdText, "Disabled"
           If ComponentsParent = Body Then
               AddPCLproperty "PD (Bottom)", .PD6, wdText, "Disabled"
           End If
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
            
       
              
        Case ArmorComplexFacing
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "TL", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Surface Coating", .coating, wdList, "Coating", "none", "reflective", "retro-reflective"
           AddPCLproperty "Radiation Shielding", .radiation, wdBool, "Radiation"
           AddPCLproperty "Thermal Superconductor", .thermal, wdBool, "Thermal"
           AddPCLproperty "Reactive Armor Plating", .rap, wdBool, "RAP"
           AddPCLproperty "Electrified", .electrified, wdBool, "Electrified"
           listarray = .FillMaterial 'only needs to be filled once for all sides
           AddPCLproperty "Material (Right)", .material1, wdList, "Material1", listarray
           AddPCLproperty "Material (Left)", .material2, wdList, "Material2", listarray
           AddPCLproperty "Material (Front)", .material3, wdList, "Material3", listarray
           AddPCLproperty "Material (Back)", .material4, wdList, "Material4", listarray
           AddPCLproperty "Material (Top)", .material5, wdList, "Material5", listarray
           If ComponentsParent = Body Then
               AddPCLproperty "Material (Bottom)", .material6, wdList, "Material6", listarray
           End If
           listarray = .FillQuality(.material1)
           AddPCLproperty "Quality (Right)", .Quality1, wdList, "Quality1", listarray
           listarray = .FillQuality(.material2)
           AddPCLproperty "Quality (Left)", .Quality2, wdList, "Quality2", listarray
           listarray = .FillQuality(.material3)
           AddPCLproperty "Quality (Front)", .Quality3, wdList, "Quality3", listarray
           listarray = .FillQuality(.material4)
           AddPCLproperty "Quality (Back)", .Quality4, wdList, "Quality4", listarray
           listarray = .FillQuality(.material5)
           AddPCLproperty "Quality (Top)", .Quality5, wdList, "Quality5", listarray
           If ComponentsParent = Body Then
                listarray = .FillQuality(.material6)
               AddPCLproperty "Quality (Bottom)", .Quality6, wdList, "Quality6", listarray
           End If
           AddPCLproperty "DR (Right)", .dr1, wdNumber, "DR1", listarray
           AddPCLproperty "DR (Left)", .dr2, wdNumber, "DR2", listarray
           AddPCLproperty "DR (Front)", .dr3, wdNumber, "DR3", listarray
           AddPCLproperty "DR (Back)", .dr4, wdNumber, "DR4", listarray
           AddPCLproperty "DR (Top)", .dr5, wdNumber, "DR5", listarray
           If ComponentsParent = Body Then
                AddPCLproperty "DR (Bottom)", .dr6, wdNumber, "DR6", listarray
           End If
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Average DR", .AverageDR, wdText, "Disabled"
           AddPCLproperty "Effective DR (Right)", .EffectiveDR1, wdText, "Disabled"
           AddPCLproperty "Effective DR (Left)", .EffectiveDR2, wdText, "Disabled"
           AddPCLproperty "Effective DR (Front)", .EffectiveDR3, wdText, "Disabled"
           AddPCLproperty "Effective DR (Back)", .EffectiveDR4, wdText, "Disabled"
           AddPCLproperty "Effective DR (Top)", .EffectiveDR5, wdText, "Disabled"
           If ComponentsParent = Body Then
               AddPCLproperty "Effective DR (Bottom)", .EffectiveDR6, wdText, "Disabled"
           End If
           AddPCLproperty "PD (Right)", .PD1, wdText, "Disabled"
           AddPCLproperty "PD (Left)", .PD2, wdText, "Disabled"
           AddPCLproperty "PD (Front)", .PD3, wdText, "Disabled"
           AddPCLproperty "PD (Back)", .PD4, wdText, "Disabled"
           AddPCLproperty "PD (Top)", .PD5, wdText, "Disabled"
           If ComponentsParent = Body Then
               AddPCLproperty "PD (Bottom)", .PD6, wdText, "Disabled"
           End If
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
        
        
        Case ArmorComponent
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           listarray = .FillMaterial
           AddPCLproperty "Material", .material, wdList, "Material", listarray
           listarray = .FillQuality(.material)
           AddPCLproperty "Quality", .Quality, wdList, "Quality", listarray
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
            AddPCLproperty "PD", .PD, wdText, "Disabled"
            AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
            AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
             
        
        Case ArmorLocation
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           listarray = .FillMaterial
           AddPCLproperty "Material", .material, wdList, "Material", listarray
           listarray = .FillQuality(.material)
           AddPCLproperty "Quality", .Quality, wdList, "Quality", listarray
           AddPCLproperty "Surface Coating", .coating, wdList, "Coating", "none", "reflective", "retro-reflective"
           AddPCLproperty "Radiation Shielding", .radiation, wdBool, "Radiation"
           AddPCLproperty "Thermal Superconductor", .thermal, wdBool, "Thermal"
           AddPCLproperty "Reactive Armor Plating", .rap, wdBool, "RAP"
           AddPCLproperty "Electrified", .electrified, wdBool, "Electrified"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
            If ComponentsParent = Body Then
                AddPCLproperty "PD (Right)", .PD1, wdText, "Disabled"
                AddPCLproperty "PD (Left)", .PD2, wdText, "Disabled"
                AddPCLproperty "PD (Front)", .PD3, wdText, "Disabled"
                AddPCLproperty "PD (Back)", .PD4, wdText, "Disabled"
                AddPCLproperty "PD (Top)", .PD5, wdText, "Disabled"
            ElseIf ComponentsParent = Turret Or ComponentsParent = Popturret Then
                AddPCLproperty "PD", .PD6, wdText, "Disabled"
            Else
                AddPCLproperty "PD", .PD, wdText, "Disabled"
            End If
            AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
            AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
            
        Case ArmorOverall, ArmorWheelGuard, ArmorGunShield
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           listarray = .FillMaterial
           AddPCLproperty "Material", .material, wdList, "Material", listarray
           listarray = .FillQuality(.material)
           AddPCLproperty "Quality", .Quality, wdList, "Quality", listarray
           AddPCLproperty "Surface Coating", .coating, wdList, "Coating", "none", "reflective", "retro-reflective"
           AddPCLproperty "Radiation Shielding", .radiation, wdBool, "Radiation"
           AddPCLproperty "Thermal Superconductor", .thermal, wdBool, "Thermal"
           AddPCLproperty "Reactive Armor Plating", .rap, wdBool, "RAP"
           AddPCLproperty "Electrified", .electrified, wdBool, "Electrified"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "PD", .PD, wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           
        Case ArmorOpenFrame
        
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           listarray = .FillMaterial
           AddPCLproperty "Material", .material, wdList, "Material", listarray
           listarray = .FillQuality(.material)
           AddPCLproperty "Quality", .Quality, wdList, "Quality", listarray
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "PD", .PD, wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
                   
    End Select
End With
End Sub

Private Sub ShowPropsForMannedVehicles(ByVal component As Integer, ByVal Key As String)
'//////////////////////////////////////////////
'Manned Vehicle Components
'//////////////////////////////////////////////
 With m_oCurrentVeh.Components(Key)

' Fill the window with properties for the correct Collection item
        
    Select Case component
        Case PrimitiveManeuverControl
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
            
            
        Case ElectronicDivingControl, _
        ComputerizedDivingControl, MechanicalManeuverControl, _
        ElectronicManeuverControl, ComputerizedManeuverControl, _
         MechanicalDivingControl
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Duplicate?", .duplicate, wdBool, "Duplicate"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
            
            
        Case BattlesuitSystem
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Pilot Weight", .PilotWeight, wdDouble, "PilotWeight"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume1", Format(.Volume1, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
        Case FormFittingBattleSuitSystem
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Pilot Weight", .PilotWeight, wdDouble, "PilotWeight"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight (w/out Pilot)", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Body Volume", Format(.Volume1, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Turret Volume", Format(.Volume2, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Arm Volume (each)", Format(.Volume3, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Leg Volume (each)", Format(.Volume4, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
        Case CrampedSeat, NormalSeat, RoomySeat
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Exposed?", .Exposed, wdBool, "Exposed"
           AddPCLproperty "G-Seat?", .GSeat, wdBool, "GSeat"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        Case CycleSeat
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
        
        Case CrampedStandingRoom, NormalStandingRoom, RoomyStandingRoom
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Exposed?", .Exposed, wdBool, "Exposed"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        Case Hammock, Bunk, SmallGalley
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Added Volume", .AddedVolume, wdDouble, "AddedVolume"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
           
        Case Cabin, LuxuryCabin, Suite, LuxurySuite
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Occupancy", .Occupancy, wdNumber, "Occupancy"
           AddPCLproperty "G-Seats?", .GSeat, wdBool, "Gseat"
           AddPCLproperty "Added Volume", .AddedVolume, wdDouble, "AddedVolume"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        Case CrampedCrewStation, NormalCrewStation, RoomyCrewStation
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Assignment", frmNotes, wdObject, "StationFunction"
           AddPCLproperty "Bridge Access Space?", .BridgeAccessSpace, wdBool, "BridgeAccessSpace"
           AddPCLproperty "Exposed?", .Exposed, wdBool, "Exposed"
           AddPCLproperty "G-Seat?", .GSeat, wdBool, "GSeat"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
        
        Case CycleCrewStation, HarnessCrewStation
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "Assignment", frmNotes, wdObject, "StationFunction"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        Case ArtificialGravityUnit
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
        Case EnvironmentalControl, NBCKit, FullLifeSystem, TotalLifeSystem
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "# People", .People, wdNumber, "People"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        
        Case LimitedLifeSystem
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "# People", .People, wdNumber, "People"
           AddPCLproperty "# Man Days", .ManDays, wdDouble, "ManDays"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
            
        Case EjectionSeat, CrewEscapeCapsule, Airbag, CrashWeb, WombTank, GravityWeb
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           If component = CrewEscapeCapsule Then
                AddPCLproperty "Max Occupancy", .Occupancy, wdNumber, "Occupancy"
           End If
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           If component = GravityWeb Then
                AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           End If
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
    
        Case GravCompensator
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           AddPCLproperty "G reduction", .GReduction, wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            
        Case Provisions
           AddPCLproperty "Settings", "", wdText, "Disabled"
           AddPCLproperty "# days worth", .occupancydays, wdNumber, "occupancydays"
           AddPCLproperty "Settings", .Setting, wdList, "Setting", "auto", "light", "heavy"
           AddPCLproperty "DR", .dr, wdNumber, "DR"
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
 
 End Select
End With
End Sub

