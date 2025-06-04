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
vbwProfiler.vbwProcIn 255
vbwProfiler.vbwExecuteLine 4810
    s = " " & s
vbwProfiler.vbwExecuteLine 4811
    formatCaption = s
vbwProfiler.vbwProcOut 255
vbwProfiler.vbwExecuteLine 4812
End Function

Function formatValue(ByVal lngDatatype As Long, ByVal lngUnitType As Long, v As Variant) As Variant
vbwProfiler.vbwProcIn 256
vbwProfiler.vbwExecuteLine 4813
    formatValue = v
vbwProfiler.vbwProcOut 256
vbwProfiler.vbwExecuteLine 4814
End Function

Sub AddPCLproperty(ByRef vValue As Variant, ByRef oPropItem As cPropertyItem)
vbwProfiler.vbwProcIn 257
    Dim lngNewIndex As Long
    Dim lngDatatype As Long
    Dim lngUnitType As Long
    Dim sCaption As String
    Dim dblValue As Double
    Dim vItem As Variant
    Dim i As Long
    'Const SEPERATOR = "___________________"
vbwProfiler.vbwExecuteLine 4815
    Const SEPERATOR = "===================="

vbwProfiler.vbwExecuteLine 4816
    lngDatatype = oPropItem.Datatype ' proplist data type (e.g. list, float, number, text, etc)
vbwProfiler.vbwExecuteLine 4817
    lngUnitType = oPropItem.UnitType ' gvd unit type used by the unit converter
vbwProfiler.vbwExecuteLine 4818
    sCaption = oPropItem.Caption

vbwProfiler.vbwExecuteLine 4819
    sCaption = formatCaption(lngDatatype, lngUnitType, sCaption)
vbwProfiler.vbwExecuteLine 4820
    vValue = formatValue(lngDatatype, lngUnitType, vValue)
    'todo: when finally done, check out for places where i handle doubles and singles together to make sure its all correct.
    ' see notes below on some things to look out for

    ' NOTE: I use double's for most final stats, however I dont for table data since it would require double the amount of space
    ' and there are ALOT of tables.  Also used in our m_UserInput() variable which gets displayed directly here in the proplist.
    ' However, since the PLC1 doesnt support a "single" datatype, I must use the PLC1's double datatype
    ' which unfortunately will show approximation errors.  To fix this, we must convert the single to a string first and then assign that to the double.
    ' NOTE: Double to single doesnt cause these errors (as long as it doesnt overflow obviously) but single to double does.  Keep this in mind during
    ' stats calculations where we might be mixing the SINGLE's in our tables with the DOUBLE's required for the final answer.
vbwProfiler.vbwExecuteLine 4821
    If lngDatatype = wdDouble Then
vbwProfiler.vbwExecuteLine 4822
        dblValue = Val(CStr(vValue))
vbwProfiler.vbwExecuteLine 4823
        frmDesigner.PLC1.AddItem dblValue, lngDatatype
'vbwLine 4824:    ElseIf lngDatatype = wdHeader Then
    ElseIf vbwProfiler.vbwExecuteLine(4824) Or lngDatatype = wdHeader Then
vbwProfiler.vbwExecuteLine 4825
        With frmDesigner.PLC1
vbwProfiler.vbwExecuteLine 4826
            .AddItem SEPERATOR, 0 ' change it back from -99 to 0 so that we can display the seperator line
vbwProfiler.vbwExecuteLine 4827
            .ItemDisabledTextBold(.NewIndex) = True
vbwProfiler.vbwExecuteLine 4828
        End With
    Else
vbwProfiler.vbwExecuteLine 4829 'B
vbwProfiler.vbwExecuteLine 4830
        frmDesigner.PLC1.AddItem vValue, lngDatatype
    End If
vbwProfiler.vbwExecuteLine 4831 'B

vbwProfiler.vbwExecuteLine 4832
    With frmDesigner.PLC1
vbwProfiler.vbwExecuteLine 4833
        lngNewIndex = .NewIndex
vbwProfiler.vbwExecuteLine 4834
        .ItemData(lngNewIndex) = lngDatatype
vbwProfiler.vbwExecuteLine 4835
        .CaptionString(lngNewIndex) = sCaption
vbwProfiler.vbwExecuteLine 4836
        .DescriptionString(lngNewIndex) = oPropItem.Notes
vbwProfiler.vbwExecuteLine 4837
        .ItemDisabled(lngNewIndex) = oPropItem.ReadOnly
vbwProfiler.vbwExecuteLine 4838
    End With

    'todo: how do we handle ammolists, guidancelists, or material/quality lists for armor?
    '      the armor layer would have to update its property item's list every time techlevel was changed.
    '      and if tech level was changed to somethign not supported by the current material, the material
    '     would default to first available at that techlevel. (e.g. cheap wood).  The only way to do this
    '     properly is use multiple matrices so that the values can be looked up based on the tech level and material.
vbwProfiler.vbwExecuteLine 4839
    If lngDatatype = wdList Then
vbwProfiler.vbwExecuteLine 4840
        vItem = oPropItem.List
vbwProfiler.vbwExecuteLine 4841
        If IsArray(vItem) Then
vbwProfiler.vbwExecuteLine 4842
            For i = LBound(vItem) To UBound(vItem)
vbwProfiler.vbwExecuteLine 4843
                If vItem(i) <> "" Then
vbwProfiler.vbwExecuteLine 4844
                    frmDesigner.PLC1.AddListItem (lngNewIndex), vItem(i)
                End If
vbwProfiler.vbwExecuteLine 4845 'B
vbwProfiler.vbwExecuteLine 4846
            Next
        End If
vbwProfiler.vbwExecuteLine 4847 'B
    End If
vbwProfiler.vbwExecuteLine 4848 'B

vbwProfiler.vbwExecuteLine 4849
    vValue = Empty
vbwProfiler.vbwProcOut 257
vbwProfiler.vbwExecuteLine 4850
End Sub

Public Sub PropertyChanged(ByVal index As Long)
vbwProfiler.vbwProcIn 258
vbwProfiler.vbwExecuteLine 4851
   Const LNG_LENGTH = 4
   Dim oProp As cPropertyItem
   Dim oNode As cINode
   Dim oDisplay As cIDisplay
   Dim lptr As Long
   Dim sClassname As String
   Dim lngInterfaceID As Long
   Dim lngRet As Long
vbwProfiler.vbwExecuteLine 4852
   Debug.Print "Object Handle = " & frmDesigner.PLC1.Tag & " PLC1 INDEX = " & index

vbwProfiler.vbwExecuteLine 4853
   On Error GoTo err
vbwProfiler.vbwExecuteLine 4854
   lptr = Val(frmDesigner.PLC1.Tag)
vbwProfiler.vbwExecuteLine 4855
   CopyMemory oNode, lptr, LNG_LENGTH

vbwProfiler.vbwExecuteLine 4856
   If Not oNode Is Nothing Then
vbwProfiler.vbwExecuteLine 4857
        sClassname = oNode.Classname
vbwProfiler.vbwExecuteLine 4858
        Set oDisplay = oNode

        ' NOTE: its imperative that the order of properties in the proplist match the order of properties in the object
vbwProfiler.vbwExecuteLine 4859
        Set oProp = oDisplay.getPropertyItemByIndex(index)
vbwProfiler.vbwExecuteLine 4860
        lngInterfaceID = oProp.interfaceid

vbwProfiler.vbwExecuteLine 4861
        If Not oProp.ReadOnly Then
            ' todo: i think this select case is "ok"... anyway to modify it or perhaps consolidate the code
            '       with modProperties:PropertiesShow() ??  Actually probably not, it uses "vbGet" and this uses
            '       vbLet.  Actually what we should do then is move this out into a seperate function
vbwProfiler.vbwExecuteLine 4862
            Debug.Print "modProperties:PropertyChanged() -- InterfaceID = " & lngInterfaceID & " PropertyName = " & oProp.CallByName
vbwProfiler.vbwExecuteLine 4863
             Select Case lngInterfaceID
'vbwLine 4864:                Case INTERFACE_COMPONENT
                Case IIf(vbwProfiler.vbwExecuteLine(4864), VBWPROFILER_EMPTY, _
        INTERFACE_COMPONENT)
                    Dim oComponent As cIComponent
vbwProfiler.vbwExecuteLine 4865
                    Set oComponent = oNode
vbwProfiler.vbwExecuteLine 4866
                    CallByName oComponent, oProp.CallByName, VbLet, frmDesigner.PLC1.value(index)

'vbwLine 4867:                Case INTERFACE_CONTAINER
                Case IIf(vbwProfiler.vbwExecuteLine(4867), VBWPROFILER_EMPTY, _
        INTERFACE_CONTAINER)
                    Dim oContainer As cIContainer
vbwProfiler.vbwExecuteLine 4868
                    Set oContainer = oNode
vbwProfiler.vbwExecuteLine 4869
                    CallByName oContainer, oProp.CallByName, VbLet, frmDesigner.PLC1.value(index)
'vbwLine 4870:                Case INTERFACE_VEHICLE_DESCRIPTION
                Case IIf(vbwProfiler.vbwExecuteLine(4870), VBWPROFILER_EMPTY, _
        INTERFACE_VEHICLE_DESCRIPTION)
                    Dim oVehicle As cVehicle
vbwProfiler.vbwExecuteLine 4871
                    Set oVehicle = oNode
vbwProfiler.vbwExecuteLine 4872
                    CallByName oVehicle.Description, oProp.CallByName, VbLet, frmDesigner.PLC1.value(index)
'vbwLine 4873:                Case INTERFACE_VEHICLE_VERSION
                Case IIf(vbwProfiler.vbwExecuteLine(4873), VBWPROFILER_EMPTY, _
        INTERFACE_VEHICLE_VERSION)
vbwProfiler.vbwExecuteLine 4874
                    CallByName oVehicle.version, oProp.CallByName, VbLet, frmDesigner.PLC1.value(index)

'vbwLine 4875:                Case INTERFACE_VEHICLE_AUTHOR
                Case IIf(vbwProfiler.vbwExecuteLine(4875), VBWPROFILER_EMPTY, _
        INTERFACE_VEHICLE_AUTHOR)
vbwProfiler.vbwExecuteLine 4876
                    CallByName oVehicle.author, oProp.CallByName, VbLet, frmDesigner.PLC1.value(index)

'vbwLine 4877:                Case INTERFACE_NODE
                Case IIf(vbwProfiler.vbwExecuteLine(4877), VBWPROFILER_EMPTY, _
        INTERFACE_NODE)
vbwProfiler.vbwExecuteLine 4878
                    CallByName oNode, oProp.CallByName, VbLet, frmDesigner.PLC1.value(index)
'vbwLine 4879:                Case INTERFACE_DISPLAY
                Case IIf(vbwProfiler.vbwExecuteLine(4879), VBWPROFILER_EMPTY, _
        INTERFACE_DISPLAY)
vbwProfiler.vbwExecuteLine 4880
                    CallByName oDisplay, oProp.CallByName, VbLet, frmDesigner.PLC1.value(index)
'vbwLine 4881:                Case INTERFACE_BUILD
                Case IIf(vbwProfiler.vbwExecuteLine(4881), VBWPROFILER_EMPTY, _
        INTERFACE_BUILD)
                    Dim oBuild As cIBuild
vbwProfiler.vbwExecuteLine 4882
                    Set oBuild = oNode
vbwProfiler.vbwExecuteLine 4883
                    If oProp.CallBytype = 1 Then
                        ' since this is a function call, we have to modify the name to append "get"
                        ' also, if its a wdList, we need to pass the subscript of the selected list item
vbwProfiler.vbwExecuteLine 4884
                        If oProp.Datatype = wdList Then
vbwProfiler.vbwExecuteLine 4885
                            lngRet = oProp.getSelectionIndexFromValue(frmDesigner.PLC1.value(index))
vbwProfiler.vbwExecuteLine 4886
                            If lngRet <> -1 Then
vbwProfiler.vbwExecuteLine 4887
                                CallByName oBuild, "set" & oProp.CallByName, VbMethod, oProp.Subscript, lngRet
                            Else
vbwProfiler.vbwExecuteLine 4888 'B
                                ' an error and there shouldnt be any
vbwProfiler.vbwExecuteLine 4889
                                MsgBox "modProperties:PropertyChanged() -- Invalid wdList subscript"
                            End If
vbwProfiler.vbwExecuteLine 4890 'B
                        ' else its userinput and we pass the actual value
'vbwLine 4891:                        ElseIf oProp.Datatype = wdDouble Then
                        ElseIf vbwProfiler.vbwExecuteLine(4891) Or oProp.Datatype = wdDouble Then
vbwProfiler.vbwExecuteLine 4892
                            CallByName oBuild, "set" & oProp.CallByName, VbMethod, oProp.Subscript, frmDesigner.PLC1.value(index)
                        Else ' we should never reach here
vbwProfiler.vbwExecuteLine 4893 'B

                        End If
vbwProfiler.vbwExecuteLine 4894 'B
                    Else
vbwProfiler.vbwExecuteLine 4895 'B
vbwProfiler.vbwExecuteLine 4896
                        CallByName oBuild, oProp.CallByName, VbLet, frmDesigner.PLC1.value(index)
                    End If
vbwProfiler.vbwExecuteLine 4897 'B
'vbwLine 4898:                Case INTERFACE_SURFACE
                Case IIf(vbwProfiler.vbwExecuteLine(4898), VBWPROFILER_EMPTY, _
        INTERFACE_SURFACE)
                    Dim oSurface As cSurface
vbwProfiler.vbwExecuteLine 4899
                    Set oSurface = oNode
vbwProfiler.vbwExecuteLine 4900
                    lngRet = oProp.getSelectionIndexFromValue(frmDesigner.PLC1.value(index))
vbwProfiler.vbwExecuteLine 4901
                    If lngRet <> -1 Then
vbwProfiler.vbwExecuteLine 4902
                        CallByName oSurface, oProp.CallByName, VbLet, lngRet
                    Else
vbwProfiler.vbwExecuteLine 4903 'B
                        ' an error and there shouldnt be any
vbwProfiler.vbwExecuteLine 4904
                        MsgBox "modProperties:PropertyChanged() -- Invalid wdList subscript"
                    End If
vbwProfiler.vbwExecuteLine 4905 'B

                Case Else
vbwProfiler.vbwExecuteLine 4906 'B
vbwProfiler.vbwExecuteLine 4907
                    Debug.Print "modProperties:PropertyChanged() -- Class Interface Not Supported."
            End Select
vbwProfiler.vbwExecuteLine 4908 'B
        Else
vbwProfiler.vbwExecuteLine 4909 'B
vbwProfiler.vbwExecuteLine 4910
             InfoPrint 1, "The property '" & oProp.Caption & "' is Read Only."
        End If
vbwProfiler.vbwExecuteLine 4911 'B

vbwProfiler.vbwExecuteLine 4912
        Set oProp = Nothing
   End If
vbwProfiler.vbwExecuteLine 4913 'B
vbwProfiler.vbwExecuteLine 4914
   CopyMemory oNode, 0&, LNG_LENGTH
vbwProfiler.vbwExecuteLine 4915
   Set oDisplay = Nothing
vbwProfiler.vbwExecuteLine 4916
   DoEvents
'    UpdateVehicle todo: all of these may be obosolete under new code base EXCEPT when user hits F5 or when forcing recalc after laoding saved vehicle

    '//place the cell back into its original spot. Todo: why is this needed?
vbwProfiler.vbwExecuteLine 4917
    frmDesigner.PLC1.ListIndex = index
vbwProfiler.vbwExecuteLine 4918
    p_bChangedFlag = True ' JAW 2000.05.07 change has been made in vehicle/component
vbwProfiler.vbwProcOut 258
vbwProfiler.vbwExecuteLine 4919
   Exit Sub
err:
vbwProfiler.vbwExecuteLine 4920
    Debug.Print "modProperties:PropertyChanged() -- Error #" & err.Number & " " & err.Description
vbwProfiler.vbwExecuteLine 4921
    If Not oNode Is Nothing Then
vbwProfiler.vbwExecuteLine 4922
        CopyMemory oNode, 0&, LNG_LENGTH
vbwProfiler.vbwExecuteLine 4923
        Set oDisplay = Nothing
vbwProfiler.vbwExecuteLine 4924
        Set oProp = Nothing
   End If
vbwProfiler.vbwExecuteLine 4925 'B

'    Todo: still need to properly handle these types.... perhaps new datatype like wdNote (signifies text greater than X chars?)
'    If frmDesigner.PLC1.DescriptionString(Index) = "StationFunction" Then
'        Load frmNotes
'        frmNotes.Tag = "crewstation"
'        frmNotes.Show vbModal, frmDesigner
'        Set frmNotes = Nothing
'

'
vbwProfiler.vbwProcOut 258
vbwProfiler.vbwExecuteLine 4926
End Sub
Public Sub Properties_Show(ByVal hNode As Long)
vbwProfiler.vbwProcIn 259
vbwProfiler.vbwExecuteLine 4927
    If hNode <= 0 Then
vbwProfiler.vbwProcOut 259
vbwProfiler.vbwExecuteLine 4928
         Exit Sub
    End If
vbwProfiler.vbwExecuteLine 4929 'B

    Dim oNode As vehicles.cINode
    Dim oDisplay As cIDisplay
    Dim oPropItem As vehicles.cPropertyItem
    Dim vValue As Variant
    Dim lngInterfaceID As Long
vbwProfiler.vbwExecuteLine 4930
    Const LNG_LENGTH = 4
    Dim index As Long
vbwProfiler.vbwExecuteLine 4931
    On Error GoTo errDefault

vbwProfiler.vbwExecuteLine 4932
    With frmDesigner.PLC1
vbwProfiler.vbwExecuteLine 4933
        .Clear
vbwProfiler.vbwExecuteLine 4934
        .ShowDescription = True
vbwProfiler.vbwExecuteLine 4935
        .Tag = hNode  'CRITICAL - needed so that when proplist attributes for a component are changed, the PLC1 code knows which item is being referenced
vbwProfiler.vbwExecuteLine 4936
    End With

vbwProfiler.vbwExecuteLine 4937
    CopyMemory oNode, hNode, LNG_LENGTH '<--- every component in the _Tree_ MUST implement cINode because thats what this pointer is for (if its not going to be rendered in the tree, it doesnt need that interface
vbwProfiler.vbwExecuteLine 4938
    Set oDisplay = oNode               '<-- every component must also obviously implement cIDisplay

vbwProfiler.vbwExecuteLine 4939
    If Not oDisplay Is Nothing Then
vbwProfiler.vbwExecuteLine 4940
    Set oPropItem = oDisplay.getfirstpropertyitem
'vbwLine 4941:        Do While Not oPropItem Is Nothing
        Do While vbwProfiler.vbwExecuteLine(4941) Or Not oPropItem Is Nothing
vbwProfiler.vbwExecuteLine 4942
            If Not oPropItem.Caption = PROPERTY_HEADER Then
vbwProfiler.vbwExecuteLine 4943
                On Error GoTo errVarName 'todo: this helps us get past bugs while under development, where properties dont exist so callbyname fails
                'NOTE: since the index location of the property in the PLC1  MUST correspond to the  index of the property in the array in our oNode
                ' we should fill in a blank line for any property that fails to properly load.

vbwProfiler.vbwExecuteLine 4944
                lngInterfaceID = oPropItem.interfaceid
vbwProfiler.vbwExecuteLine 4945
                Debug.Print "modProperties:Properties_Show() -- InterfaceID = " & lngInterfaceID & " PropertyName = " & oPropItem.CallByName
vbwProfiler.vbwExecuteLine 4946
                Select Case lngInterfaceID
                'todo: this entire select case needs to be moved to  a seperate function
                'todo: and wouldnt it just be better to call oNode.ClassName and then do the select case by TypeName??
                '      actually, I dont think I can since with composite objects (like armor inside a component) there will exist
                '      multiple interface ID's.  Using typename will not allow us to switch between interfaces.
                '      The real question is, is there a way to get rid of having a huge select case statement?
'vbwLine 4947:                    Case INTERFACE_NODE
                    Case IIf(vbwProfiler.vbwExecuteLine(4947), VBWPROFILER_EMPTY, _
        INTERFACE_NODE)
vbwProfiler.vbwExecuteLine 4948
                        vValue = CallByName(oNode, oPropItem.CallByName, VbGet)

'vbwLine 4949:                    Case INTERFACE_COMPONENT
                    Case IIf(vbwProfiler.vbwExecuteLine(4949), VBWPROFILER_EMPTY, _
        INTERFACE_COMPONENT)
                        ' NOTE: keep this twoard top of select case since its an often used case
                        Dim oComponent As cIComponent
vbwProfiler.vbwExecuteLine 4950
                        Set oComponent = oNode
vbwProfiler.vbwExecuteLine 4951
                        vValue = CallByName(oComponent, oPropItem.CallByName, VbGet)

'vbwLine 4952:                    Case INTERFACE_ARMOR
                    Case IIf(vbwProfiler.vbwExecuteLine(4952), VBWPROFILER_EMPTY, _
        INTERFACE_ARMOR)
                        Dim oArmor As cArmor
vbwProfiler.vbwExecuteLine 4953
                        Set oArmor = oNode
vbwProfiler.vbwExecuteLine 4954
                        vValue = CallByName(oArmor, oPropItem.CallByName, VbGet)

'vbwLine 4955:                    Case INTERFACE_CONTAINER
                    Case IIf(vbwProfiler.vbwExecuteLine(4955), VBWPROFILER_EMPTY, _
        INTERFACE_CONTAINER)
                        Dim oContainer As cIContainer
vbwProfiler.vbwExecuteLine 4956
                        Set oContainer = oNode
vbwProfiler.vbwExecuteLine 4957
                        vValue = CallByName(oContainer, oPropItem.CallByName, VbGet)
'vbwLine 4958:                    Case INTERFACE_VEHICLE_DESCRIPTION
                    Case IIf(vbwProfiler.vbwExecuteLine(4958), VBWPROFILER_EMPTY, _
        INTERFACE_VEHICLE_DESCRIPTION)
                        Dim oVehicle As cVehicle
vbwProfiler.vbwExecuteLine 4959
                        Set oVehicle = oNode
vbwProfiler.vbwExecuteLine 4960
                        vValue = CallByName(oVehicle.Description, oPropItem.CallByName, VbGet)

'vbwLine 4961:                    Case INTERFACE_VEHICLE_VERSION
                    Case IIf(vbwProfiler.vbwExecuteLine(4961), VBWPROFILER_EMPTY, _
        INTERFACE_VEHICLE_VERSION)
                       ' Dim oVehicle As cVehicle
vbwProfiler.vbwExecuteLine 4962
                        Set oVehicle = oNode
vbwProfiler.vbwExecuteLine 4963
                        vValue = CallByName(oVehicle.version, oPropItem.CallByName, VbGet)

'vbwLine 4964:                    Case INTERFACE_VEHICLE_AUTHOR
                    Case IIf(vbwProfiler.vbwExecuteLine(4964), VBWPROFILER_EMPTY, _
        INTERFACE_VEHICLE_AUTHOR)
                       ' Dim oVehicle As cVehicle
vbwProfiler.vbwExecuteLine 4965
                        Set oVehicle = oNode
vbwProfiler.vbwExecuteLine 4966
                        vValue = CallByName(oVehicle.author, oPropItem.CallByName, VbGet)

'vbwLine 4967:                    Case INTERFACE_BUILD
                    Case IIf(vbwProfiler.vbwExecuteLine(4967), VBWPROFILER_EMPTY, _
        INTERFACE_BUILD)
                        Dim oBuild As cIBuild
vbwProfiler.vbwExecuteLine 4968
                        Set oBuild = oNode
vbwProfiler.vbwExecuteLine 4969
                        If oPropItem.CallBytype = 1 Then
vbwProfiler.vbwExecuteLine 4970
                            vValue = CallByName(oBuild, "get" & oPropItem.CallByName, VbMethod, oPropItem.Subscript)
                            ' if its user input, we display the value
vbwProfiler.vbwExecuteLine 4971
                            If oPropItem.Datatype = wdDouble Then
                                'vValue = vValue
                            ' else its an option and we use the returned index value to find the string represenation for the selection
'vbwLine 4972:                            ElseIf oPropItem.Datatype = wdList Then
                            ElseIf vbwProfiler.vbwExecuteLine(4972) Or oPropItem.Datatype = wdList Then
vbwProfiler.vbwExecuteLine 4973
                                vValue = oPropItem.ListItem(vValue)
                            Else
vbwProfiler.vbwExecuteLine 4974 'B
vbwProfiler.vbwExecuteLine 4975
                                MsgBox "modProperties.Properties_Show() -- Error: Undefined property type."
                            End If
vbwProfiler.vbwExecuteLine 4976 'B
                        Else
vbwProfiler.vbwExecuteLine 4977 'B
vbwProfiler.vbwExecuteLine 4978
                            vValue = CallByName(oBuild, oPropItem.CallByName, VbGet)
                        End If
vbwProfiler.vbwExecuteLine 4979 'B

'vbwLine 4980:                    Case INTERFACE_SURFACE
                    Case IIf(vbwProfiler.vbwExecuteLine(4980), VBWPROFILER_EMPTY, _
        INTERFACE_SURFACE)
                        Dim oSurface As cSurface
vbwProfiler.vbwExecuteLine 4981
                        Set oSurface = oNode

                        'todo: im potentially going to wind up with the same style If/else block for every interface that has a wdList
                        '      is there another way to design this?  Well, maybe its not too many interfaces afterall?  We will see
vbwProfiler.vbwExecuteLine 4982
                        If oPropItem.Datatype = wdList Then
vbwProfiler.vbwExecuteLine 4983
                            vValue = CallByName(oSurface, oPropItem.CallByName, VbGet)
vbwProfiler.vbwExecuteLine 4984
                            vValue = oPropItem.ListItem(vValue)
                        Else
vbwProfiler.vbwExecuteLine 4985 'B
vbwProfiler.vbwExecuteLine 4986
                            vValue = CallByName(oSurface, oPropItem.CallByName, VbGet)
                        End If
vbwProfiler.vbwExecuteLine 4987 'B

                    Case Else
vbwProfiler.vbwExecuteLine 4988 'B
                        'problem
vbwProfiler.vbwExecuteLine 4989
                        InfoPrint 1, "modProperties:Properties_Show() -- Unsupported Class Interface ID '" & lngInterfaceID & "'  Cannot list property '" & oPropItem.Caption & "'"
                End Select
vbwProfiler.vbwExecuteLine 4990 'B
            End If
vbwProfiler.vbwExecuteLine 4991 'B

vbwProfiler.vbwExecuteLine 4992
            On Error GoTo errDefault
            'NOTE: this property must get added here regardless of whether there was a problem accessing its value
vbwProfiler.vbwExecuteLine 4993
            AddPCLproperty vValue, oPropItem
            ' get the next one
vbwProfiler.vbwExecuteLine 4994
            Set oPropItem = oDisplay.getnextpropertyitem
vbwProfiler.vbwExecuteLine 4995
        Loop
vbwProfiler.vbwExecuteLine 4996
        CopyMemory oNode, 0&, LNG_LENGTH
vbwProfiler.vbwExecuteLine 4997
        Set oPropItem = Nothing
vbwProfiler.vbwExecuteLine 4998
        Set oDisplay = Nothing
    End If
vbwProfiler.vbwExecuteLine 4999 'B

    'set the column width of the proplist to always be half the total width
    'todo: this should be done on event when the width of this control changes
vbwProfiler.vbwExecuteLine 5000
   frmDesigner.PLC1.ColumnWidth = frmDesigner.PLC1.Width / (2 * Screen.TwipsPerPixelX)
vbwProfiler.vbwProcOut 259
vbwProfiler.vbwExecuteLine 5001
    Exit Sub

errVarName:
vbwProfiler.vbwExecuteLine 5002
    Debug.Print "modProperties.Properties_Show() -- Could not get value for Variable Name '" & oPropItem.CallByName & "'"
vbwProfiler.vbwExecuteLine 5003
    Resume Next
errDefault:
vbwProfiler.vbwExecuteLine 5004
    Debug.Print "modProperties:Properties_Show -- Error #" & err.Number & " " & err.Description
vbwProfiler.vbwExecuteLine 5005
    If Not oNode Is Nothing Then
vbwProfiler.vbwExecuteLine 5006
        CopyMemory oNode, 0&, LNG_LENGTH
    End If
vbwProfiler.vbwExecuteLine 5007 'B
vbwProfiler.vbwExecuteLine 5008
    Set oPropItem = Nothing
vbwProfiler.vbwExecuteLine 5009
    Set oDisplay = Nothing
vbwProfiler.vbwProcOut 259
vbwProfiler.vbwExecuteLine 5010
End Sub

Public Sub DisplayPrintOutput()
vbwProfiler.vbwProcIn 260

    Dim sKey As String
    Dim lngDatatype As Long
    Dim sParentKey As String
    'todo: This could should actually be integrated with the Show_Properties since it has the same
    '      task of determine what type of node we're dealing with
vbwProfiler.vbwExecuteLine 5011
    sKey = p_ActiveNode.Key
vbwProfiler.vbwExecuteLine 5012
    lngDatatype = p_ActiveNode.Datatype
vbwProfiler.vbwExecuteLine 5013
    sParentKey = p_ActiveNode.Parent
    'display the print output in the status bar for this node
vbwProfiler.vbwExecuteLine 5014
    Select Case lngDatatype
'vbwLine 5015:        Case Body
        Case IIf(vbwProfiler.vbwExecuteLine(5015), VBWPROFILER_EMPTY, _
        Body)
            'frmDesigner.StatusBar1.Panels(1).text = m_oCurrentVeh.Body.PrintOutput
'vbwLine 5016:        Case VEHICLE_NODE
        Case IIf(vbwProfiler.vbwExecuteLine(5016), VBWPROFILER_EMPTY, _
        VEHICLE_NODE)
            ' do nothing
        ' is a component
'vbwLine 5017:        Case PERFORMANCE_NODE
        Case IIf(vbwProfiler.vbwExecuteLine(5017), VBWPROFILER_EMPTY, _
        PERFORMANCE_NODE)
'vbwLine 5018:        Case CREW_NODE
        Case IIf(vbwProfiler.vbwExecuteLine(5018), VBWPROFILER_EMPTY, _
        CREW_NODE)
'vbwLine 5019:        Case POWERSYSTEMS_NODE
        Case IIf(vbwProfiler.vbwExecuteLine(5019), VBWPROFILER_EMPTY, _
        POWERSYSTEMS_NODE)
vbwProfiler.vbwExecuteLine 5020
            If sParentKey = POWERSYSTEMS_KEY Then
                'frmDesigner.StatusBar1.Panels(1).Text = m_oCurrentVeh.Profiles(sKey).Description
            End If
vbwProfiler.vbwExecuteLine 5021 'B
'vbwLine 5022:        Case WEAPON_LINKS_NODE
        Case IIf(vbwProfiler.vbwExecuteLine(5022), VBWPROFILER_EMPTY, _
        WEAPON_LINKS_NODE)
vbwProfiler.vbwExecuteLine 5023
            If sParentKey = WEAPON_LINKS_KEY Then
                'frmDesigner.StatusBar1.Panels(1).Text = m_oCurrentVeh.WeaponProfiles(sKey).Description & "  --  " & Format(m_oCurrentVeh.WeaponProfiles(sKey).Cost, vbCurrency)
            End If
vbwProfiler.vbwExecuteLine 5024 'B
'vbwLine 5025:        Case FUELSYSTEMS_NODE
        Case IIf(vbwProfiler.vbwExecuteLine(5025), VBWPROFILER_EMPTY, _
        FUELSYSTEMS_NODE)
vbwProfiler.vbwExecuteLine 5026
            If sParentKey = FUELSYSTEMS_KEY Then
                'frmDesigner.StatusBar1.Panels(1).Text = m_oCurrentVeh.Profiles(sKey).Description
            End If
vbwProfiler.vbwExecuteLine 5027 'B
'vbwLine 5028:        Case PERFORMANCEAIR To PERFORMANCESPACE
        Case IIf(vbwProfiler.vbwExecuteLine(5028), VBWPROFILER_EMPTY, _
        PERFORMANCEAIR) To PERFORMANCESPACE
            'frmDesigner.StatusBar1.Panels(1).Text = m_oCurrentVeh.PerformanceProfiles(sKey).Description

        Case Else
vbwProfiler.vbwExecuteLine 5029 'B
            ' regular components
            'todo: uncomment
           ' frmDesigner.StatusBar1.Panels(1).text = m_oCurrentVeh.Components(sKey).PrintOutput
    End Select
vbwProfiler.vbwExecuteLine 5030 'B
vbwProfiler.vbwProcOut 260
vbwProfiler.vbwExecuteLine 5031
End Sub

Private Sub ShowPropsForPerformanceProfile()
vbwProfiler.vbwProcIn 261

vbwProfiler.vbwExecuteLine 5032
    PopulateCheckList m_oCurrentVeh.ActiveCheckListType
    Dim sKey As String
vbwProfiler.vbwExecuteLine 5033
    sKey = m_oCurrentVeh.ActiveCheckList
vbwProfiler.vbwExecuteLine 5034
    If sKey <> "" Then
vbwProfiler.vbwExecuteLine 5035
        If m_oCurrentVeh.ActiveCheckListType = WEAPON_CHECKLIST Then
vbwProfiler.vbwExecuteLine 5036
            ShowPropsForWeaponLink sKey
        Else
vbwProfiler.vbwExecuteLine 5037 'B
vbwProfiler.vbwExecuteLine 5038
            ShowPropsForPerformance sKey
        End If
vbwProfiler.vbwExecuteLine 5039 'B
    End If
vbwProfiler.vbwExecuteLine 5040 'B

vbwProfiler.vbwProcOut 261
vbwProfiler.vbwExecuteLine 5041
End Sub
Private Sub ShowPropsForPowerProfile()
vbwProfiler.vbwProcIn 262
    Dim sKey As String
vbwProfiler.vbwExecuteLine 5042
    sKey = m_oCurrentVeh.ActiveProfile
vbwProfiler.vbwExecuteLine 5043
    If sKey <> "" Then
vbwProfiler.vbwExecuteLine 5044
        m_oCurrentVeh.Profiles(sKey).Show

vbwProfiler.vbwExecuteLine 5045
        If m_oCurrentVeh.ActiveProfiletype = FUEL_PROFILE Then
vbwProfiler.vbwExecuteLine 5046
            Call ShowLinks(FUEL_PROFILE)
        Else
vbwProfiler.vbwExecuteLine 5047 'B
vbwProfiler.vbwExecuteLine 5048
            Call ShowLinks(POWER_PROFILE)
        End If
vbwProfiler.vbwExecuteLine 5049 'B
    End If
vbwProfiler.vbwExecuteLine 5050 'B
vbwProfiler.vbwProcOut 262
vbwProfiler.vbwExecuteLine 5051
End Sub
Private Sub ShowPropsForDescription()
    'these settings are available off of vehicle node
vbwProfiler.vbwProcIn 263
vbwProfiler.vbwProcOut 263
vbwProfiler.vbwExecuteLine 5052
End Sub

Private Sub ShowPropsForOptions()
    'dont need this... this shows up for vehicle node
vbwProfiler.vbwProcIn 264
vbwProfiler.vbwProcOut 264
vbwProfiler.vbwExecuteLine 5053
End Sub

Private Sub ShowPropsForCrew()
vbwProfiler.vbwProcIn 265

vbwProfiler.vbwExecuteLine 5054
    With m_oCurrentVeh.crew
vbwProfiler.vbwExecuteLine 5055
      AddPCLproperty "Settings", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5056
      AddPCLproperty "Use Recommended Crew", .UseRecommendedCrew, wdBool, "UseRecommendedCrew"
vbwProfiler.vbwExecuteLine 5057
      AddPCLproperty "Occupancy", .Occupancy, wdList, "Occupancy", "short", "long"
vbwProfiler.vbwExecuteLine 5058
      AddPCLproperty "Number of Shifts", .numshifts, wdList, "NumShifts", 1, 2, 3, 4, 5, 6, 7, 8, 9
vbwProfiler.vbwExecuteLine 5059
      AddPCLproperty "Military Vehicle", .MilitaryVehicle, wdBool, "MilitaryVehicle"

vbwProfiler.vbwExecuteLine 5060
      AddPCLproperty "Crew Quantities", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5061
      AddPCLproperty "Captains", .numcaptains, wdNumber, "NumCaptains"
vbwProfiler.vbwExecuteLine 5062
      AddPCLproperty "Officers", .NumOfficers, wdNumber, "NumOfficers"
vbwProfiler.vbwExecuteLine 5063
      AddPCLproperty "Crew Station Operators", .NumCrewStationOperators, wdNumber, "NumCrewStationOperators"
vbwProfiler.vbwExecuteLine 5064
      AddPCLproperty "Weapon Loaders", .NumWeaponLoaders, wdNumber, "NumWeaponLoaders"

vbwProfiler.vbwExecuteLine 5065
      AddPCLproperty "Rowers", .NumRowers, wdNumber, "NumRowers"
vbwProfiler.vbwExecuteLine 5066
      AddPCLproperty "Sailors", .NumSailors, wdNumber, "NumSailors"
vbwProfiler.vbwExecuteLine 5067
      AddPCLproperty "Riggers", .NumRiggers, wdNumber, "NumRiggers"
vbwProfiler.vbwExecuteLine 5068
      AddPCLproperty "Fuel Stokers", .NumFuelStokers, wdNumber, "NumFuelStokers"
vbwProfiler.vbwExecuteLine 5069
      AddPCLproperty "Mechanics", .NumMechanics, wdNumber, "NumMechanics"
vbwProfiler.vbwExecuteLine 5070
      AddPCLproperty "Service Crewmen", .NumServiceCrewmen, wdNumber, "NumServiceCrewmen"
vbwProfiler.vbwExecuteLine 5071
      AddPCLproperty "Medics", .NumMedics, wdNumber, "NumMedics"
vbwProfiler.vbwExecuteLine 5072
      AddPCLproperty "Scientists", .NumScientists, wdNumber, "NumScientists"
vbwProfiler.vbwExecuteLine 5073
      AddPCLproperty "Auxiliary Vehicle Crew", .NumAuxiliaryVehicleCrew, wdNumber, "NumAuxiliaryVehicleCrew"
vbwProfiler.vbwExecuteLine 5074
      AddPCLproperty "Stewards", .NumStewards, wdNumber, "NumStewards"

vbwProfiler.vbwExecuteLine 5075
      AddPCLproperty "Passenger Quantities", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5076
      AddPCLproperty "Luxury", .NumLuxury, wdNumber, "NumLuxury"
vbwProfiler.vbwExecuteLine 5077
      AddPCLproperty "First Class", .NumFirstClass, wdNumber, "NumFirstClass"
vbwProfiler.vbwExecuteLine 5078
      AddPCLproperty "Second Class", .NumSecondClass, wdNumber, "NumSecondClass"
vbwProfiler.vbwExecuteLine 5079
      AddPCLproperty "Steerage", .NumSteerage, wdNumber, "NumSteerage"

vbwProfiler.vbwExecuteLine 5080
      AddPCLproperty "Stats", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5081
      AddPCLproperty "Total Crew + Passengers", .TotalNumberCrewPassengers, wdNumber, "TotalNumberCrewPassengers"
vbwProfiler.vbwExecuteLine 5082
    End With
vbwProfiler.vbwProcOut 265
vbwProfiler.vbwExecuteLine 5083
End Sub
 
Private Sub ShowPropsForSurface()
vbwProfiler.vbwProcIn 266

vbwProfiler.vbwProcOut 266
vbwProfiler.vbwExecuteLine 5084
End Sub

Private Sub ShowPropsForStats()
vbwProfiler.vbwProcIn 267

vbwProfiler.vbwExecuteLine 5085
With m_oCurrentVeh.Description
    Dim vCategories() As Variant
    Dim vSubCategories() As Variant
vbwProfiler.vbwExecuteLine 5086
    Call LoadCategories(vCategories)
vbwProfiler.vbwExecuteLine 5087
    Call LoadSubCategories("Wheeled", vSubCategories)

vbwProfiler.vbwExecuteLine 5088
    On Error Resume Next '<-- todo: this is because when you first create a vehicle, the two vCategories() and vSubCategories lines that follow will raise an error
    'Vehicle Description and Authoring Information
vbwProfiler.vbwExecuteLine 5089
    AddPCLproperty "Vehicle Description", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5090
    AddPCLproperty "Name", .NickName, wdText, "NickName"
vbwProfiler.vbwExecuteLine 5091
    AddPCLproperty "Class", .Classname, wdText, "ClassName"
    'todo: Hrm... categories suck.  They were implemented originally with the intent that they could be used
    ' to filter submissions to the website.  However, i think a better way to handle this is for the user to
    ' select the cat/subcat WHEN they want to upload it.  Primarily because these categories can change
    ' on the website, and users wont always have proper categories in GVD
vbwProfiler.vbwExecuteLine 5092
    AddPCLproperty "Category", .Category, wdList, "Category", vCategories()  'todo: this one (see above)
vbwProfiler.vbwExecuteLine 5093
    AddPCLproperty "Sub Category", .subcategory, wdList, "subcategory", vSubCategories() 'todo: and this one (see above)

vbwProfiler.vbwExecuteLine 5094
    AddPCLproperty "Description", .VehicleDescription, wdText, "VehicleDescription"
vbwProfiler.vbwExecuteLine 5095
    AddPCLproperty "Details", .Details, wdText, "Details"
vbwProfiler.vbwExecuteLine 5096
    AddPCLproperty "Vision", .Vision, wdText, "Vision"
vbwProfiler.vbwExecuteLine 5097
    AddPCLproperty "Header", .Header, wdText, "Header"
vbwProfiler.vbwExecuteLine 5098
    AddPCLproperty "Footer", .Footer, wdText, "Footer"
vbwProfiler.vbwExecuteLine 5099
    AddPCLproperty "VehicleImageFileName", .VehicleImageFileName, wdText, "VehicleImageFileName"

vbwProfiler.vbwExecuteLine 5100
    AddPCLproperty "Versioning", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5101
    AddPCLproperty "Auto Increment Version", .blnAutoIncrement, wdBool, "blnAutoIncrement"
vbwProfiler.vbwExecuteLine 5102
    AddPCLproperty "Vehicle Version", .version, wdText, "version" 'todo: make sure its read only


vbwProfiler.vbwExecuteLine 5103
    AddPCLproperty "Author Info", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5104
    AddPCLproperty "Name", .author, wdText, "Author"
vbwProfiler.vbwExecuteLine 5105
    AddPCLproperty "Email", .email, wdText, "Email"
vbwProfiler.vbwExecuteLine 5106
    AddPCLproperty "Website", .url, wdText, "Url"
vbwProfiler.vbwExecuteLine 5107
End With

vbwProfiler.vbwExecuteLine 5108
 With m_oCurrentVeh.Options

vbwProfiler.vbwExecuteLine 5109
      AddPCLproperty "Miscellaneous", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5110
      AddPCLproperty "Vehicle Crafstmanship", .Quality, wdList, "Quality", "standard", "cheap", "fine", "very fine"
vbwProfiler.vbwExecuteLine 5111
      AddPCLproperty "RollStabilizers", .RollStabilizers, wdBool, "RollStabilizers"
vbwProfiler.vbwExecuteLine 5112
      AddPCLproperty "Convertible", .Convertible, wdList, "Convertible", "none", "hardtop", "ragtop"
vbwProfiler.vbwExecuteLine 5113
      AddPCLproperty "UseHardpointMountedWeights", .UseHardpointMountedWeights, wdBool, "UseHardpointMountedWeights"

vbwProfiler.vbwExecuteLine 5114
      AddPCLproperty "Payload Settings", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5115
      AddPCLproperty "Use Default Weights", .RecommendedPayload, wdBool, "RecommendedPayload"
vbwProfiler.vbwExecuteLine 5116
      AddPCLproperty "Per Person Weight", .PerPersonWeight, wdNumber, "PerPersonWeight"
vbwProfiler.vbwExecuteLine 5117
      AddPCLproperty "Per Cargo Weight", .PerCargoWeight, wdNumber, "PerCargoWeight"

vbwProfiler.vbwExecuteLine 5118
      AddPCLproperty "Access Space", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5119
      AddPCLproperty "Use Recommended Modifier?", .RecommendedAccessSpace, wdBool, "RecommendedAccessSpace"
vbwProfiler.vbwExecuteLine 5120
      AddPCLproperty "Volume Modifier", .AccessSpaceVolumeMod, wdList, "AccessSpaceVolumeMod", 0, 0.25, 0.5, 0.75, 1, 1.25, 1.5, 1.75, 2

vbwProfiler.vbwExecuteLine 5121
      AddPCLproperty "Attachments", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5122
      AddPCLproperty "Pin", .Pin, wdList, "Pin", "none", "standard", "Explosive"
vbwProfiler.vbwExecuteLine 5123
      AddPCLproperty "Ram", .Ram, wdBool, "ram"
vbwProfiler.vbwExecuteLine 5124
      AddPCLproperty "Bulldozer", .Bulldozer, wdBool, "bulldozer"
vbwProfiler.vbwExecuteLine 5125
      AddPCLproperty "Plow", .Plow, wdBool, "Plow"
vbwProfiler.vbwExecuteLine 5126
      AddPCLproperty "Hitch", .Hitch, wdBool, "hitch"
vbwProfiler.vbwExecuteLine 5127
    End With

vbwProfiler.vbwExecuteLine 5128
 With m_oCurrentVeh.surface
vbwProfiler.vbwExecuteLine 5129
       AddPCLproperty "Hull and Hydro Options", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5130
       AddPCLproperty "Streamlining", .StreamLining, wdList, "Streamlining", "none", "fair", "good", "very good", "superior", "excellent", "radical"
vbwProfiler.vbwExecuteLine 5131
       AddPCLproperty "Floatation Hull", .FloatationHull, wdBool, "floatationhull"
vbwProfiler.vbwExecuteLine 5132
       AddPCLproperty "Submerisible Hull (TL5)", .Submersible, wdBool, "Submersible"
vbwProfiler.vbwExecuteLine 5133
       AddPCLproperty "Hydrodynamic Lines", .HydrodynamicLines, wdList, "Hydrodynamiclines", "none", "mediocre", "average", "fine", "very fine", "submarine"
       'AddPCLproperty "Roll Stabilizers (TL7)", .RollStabilizers, wdBool, "rollstabilizers"
vbwProfiler.vbwExecuteLine 5134
       AddPCLproperty "Waterproof", .WaterProof, wdBool, "waterproof"
vbwProfiler.vbwExecuteLine 5135
       AddPCLproperty "Sealed (TL5)", .Sealed, wdBool, "Sealed"
vbwProfiler.vbwExecuteLine 5136
       AddPCLproperty "Cata/Tri(maran)", .CataTrimaran, wdList, "catatrimaran", "none", "catamaran", "trimaran"



vbwProfiler.vbwExecuteLine 5137
       AddPCLproperty "Concealment", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5138
       AddPCLproperty "Camouflage", .Camouflage, wdBool, "Camouflage"
vbwProfiler.vbwExecuteLine 5139
       AddPCLproperty "Infrared Cloaking (TL7)", .infraredcloaking, wdList, "InfraredCloaking", "none", "basic", "radical"
vbwProfiler.vbwExecuteLine 5140
       AddPCLproperty "Emission Cloaking (TL8)", .EmissionCloaking, wdList, "EmissionCloaking", "none", "basic", "radical"
vbwProfiler.vbwExecuteLine 5141
       AddPCLproperty "Sound Baffling (TL7)", .SoundBaffling, wdList, "SoundBaffling", "none", "basic", "radical"
vbwProfiler.vbwExecuteLine 5142
       AddPCLproperty "Stealth (TL7)", .Stealth, wdList, "Stealth", "none", "basic", "radical"
vbwProfiler.vbwExecuteLine 5143
       AddPCLproperty "Liquid Crystal Skin (TL8)", .LiquidCrystal, wdBool, "LiquidCrystal"
vbwProfiler.vbwExecuteLine 5144
       AddPCLproperty "PsiShielding (TL8)", .PsiShielding, wdBool, "PsiShielding"
vbwProfiler.vbwExecuteLine 5145
       AddPCLproperty "Chameleon", .Chameleon, wdList, "Chameleon", "none", "basic", "instant", "intruder"

vbwProfiler.vbwExecuteLine 5146
       AddPCLproperty "Magic Levitation", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5147
       AddPCLproperty "Enabled", .bMagicLevitation, wdBool, "bMagicLevitation"
vbwProfiler.vbwExecuteLine 5148
       AddPCLproperty "Energy Cost Per Pound", .MagicLevitationEnergyCostPerPound, wdDouble, "MagicLevitationEnergyCostPerPound"

vbwProfiler.vbwExecuteLine 5149
       AddPCLproperty "Antigravity Coating", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5150
       AddPCLproperty "Enabled", .bAntigravityCoating, wdBool, "bAntigravityCoating"
vbwProfiler.vbwExecuteLine 5151
       AddPCLproperty "Cost Per Sq ft", .AntigravityCoatingCostPerSquareFoot, wdDouble, "AntigravityCoatingCostPerSquareFoot"
vbwProfiler.vbwExecuteLine 5152
       AddPCLproperty "Surface Area Useage", .AntigravityCoatingSurfaceAreaUseage, wdList, "AntigravityCoatingSurfaceAreaUseage", "Body", "Vehicle"

vbwProfiler.vbwExecuteLine 5153
       AddPCLproperty "Super Science Coating", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5154
       AddPCLproperty "Enabled", .bSuperScienceCoating, wdBool, "bSuperScienceCoating"
vbwProfiler.vbwExecuteLine 5155
       AddPCLproperty "Cost Per Sq ft", .SuperScienceCoatingCostPerSquareFoot, wdDouble, "SuperScienceCoatingCostPerSquareFoot"
vbwProfiler.vbwExecuteLine 5156
       AddPCLproperty "Surface Area Useage", .SuperScienceCoatingSurfaceAreaUseage, wdList, "SuperScienceCoatingSurfaceAreaUseage", "Body", "Vehicle"

vbwProfiler.vbwExecuteLine 5157
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


vbwProfiler.vbwProcOut 267
vbwProfiler.vbwExecuteLine 5158
End Sub

Public Sub ShowPropsForPerformance(ByVal Key As String)
vbwProfiler.vbwProcIn 268

    'PerformanceType is stored in  m_oCurrentVeh.PerformanceProfiles(Key).Datatype

vbwProfiler.vbwExecuteLine 5159
    With m_oCurrentVeh.PerformanceProfiles(Key)

vbwProfiler.vbwExecuteLine 5160
        Select Case .Datatype
        'JAW 2000.06.18
        'Added takeoff/land

'vbwLine 5161:            Case PERFORMANCELEG
            Case IIf(vbwProfiler.vbwExecuteLine(5161), VBWPROFILER_EMPTY, _
        PERFORMANCELEG)
vbwProfiler.vbwExecuteLine 5162
                AddPCLproperty "Legged Performance", "", wdText, PROPERTY_HEADER

vbwProfiler.vbwExecuteLine 5163
                AddPCLproperty "Thrust Options", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5164
                AddPCLproperty "PercentThrust", .percentthrust, wdNumber, "PercentThrust"
vbwProfiler.vbwExecuteLine 5165
                AddPCLproperty "TreatTiltRotorsAsPropellers", .TreatTiltRotorsAsPropellers, wdBool, "TreatTiltRotorsAsPropellers"
vbwProfiler.vbwExecuteLine 5166
                AddPCLproperty "AfterBurnersOn", .AfterBurnersOn, wdBool, "AfterBurnersOn"


vbwProfiler.vbwExecuteLine 5167
                AddPCLproperty "Streamlining", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5168
                AddPCLproperty "HardPointsOn", .HardPointsOn, wdBool, "HardPointsOn"
vbwProfiler.vbwExecuteLine 5169
                AddPCLproperty "WheelsSkidsExtended", .WheelsSkidsExtended, wdBool, "WheelsSkidsExtended"
vbwProfiler.vbwExecuteLine 5170
                AddPCLproperty "PopTurretsExtended", .PopTurretsExtended, wdBool, "PopTurretsExtended"

vbwProfiler.vbwExecuteLine 5171
                AddPCLproperty "Weight Percentages", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5172
                AddPCLproperty "% Crew", .PercentCrewWeight, wdNumber, "PercentCrewWeight"
vbwProfiler.vbwExecuteLine 5173
                AddPCLproperty "% Fuel", .PercentFuelWeight, wdNumber, "PercentFuelWeight"
vbwProfiler.vbwExecuteLine 5174
                AddPCLproperty "% Cargo", .PercentCargoWeight, wdNumber, "PercentCargoWeight"
vbwProfiler.vbwExecuteLine 5175
                AddPCLproperty "% Hardpoints/Bays Load", .PercentHardpointWeight, wdNumber, "PercentHardpointWeight"
vbwProfiler.vbwExecuteLine 5176
                AddPCLproperty "% Provisions", .PercentProvisionWeight, wdNumber, "PercentProvisionWeight"
vbwProfiler.vbwExecuteLine 5177
                AddPCLproperty "% Ammunitions", .PercentAmmunitionWeight, wdNumber, "PercentAmmunitionWeight"
vbwProfiler.vbwExecuteLine 5178
                AddPCLproperty "% PercentAuxVehicleWeight", .PercentAuxVehicleWeight, wdNumber, "PercentAuxVehicleWeight"

vbwProfiler.vbwExecuteLine 5179
                AddPCLproperty "Statistics", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5180
                AddPCLproperty "Total Drivetrain Power", Format(.gtotalmotivepower, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5181
                AddPCLproperty "gSpeed", Format(.gTopSpeed, "standard") & " mph", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5182
                AddPCLproperty "gOffRd", Format(.gOffRoad, "standard") & " mph", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5183
                AddPCLproperty "gAccel", .gAcceleration & " mph/s", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5184
                AddPCLproperty "gDecel", .gDeceleration & " mph/s", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5185
                AddPCLproperty "gSR", .gStability, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5186
                AddPCLproperty "gMR", .gManeuverability, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5187
                AddPCLproperty "gP", Format(.gPressure, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5188
                AddPCLproperty "gPDescr", .gPressureDescription, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5189
                AddPCLproperty "Advisory", .Advisory, wdText, "Disabled"

'vbwLine 5190:            Case PERFORMANCETRACK
            Case IIf(vbwProfiler.vbwExecuteLine(5190), VBWPROFILER_EMPTY, _
        PERFORMANCETRACK)
vbwProfiler.vbwExecuteLine 5191
                AddPCLproperty "Tracked Performance", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5192
                AddPCLproperty "gSpeed", Format(.gTopSpeed, "standard") & " mph", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5193
                AddPCLproperty "gOffRd", Format(.gOffRoad, "standard") & " mph", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5194
                AddPCLproperty "gAccel", .gAcceleration & " mph/s", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5195
                AddPCLproperty "gDecel", .gDeceleration & " mph/s", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5196
                AddPCLproperty "gSR", .gStability, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5197
                AddPCLproperty "gMR", .gManeuverability, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5198
                AddPCLproperty "gP", Format(.gPressure, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5199
                AddPCLproperty "gPDescr", .gPressureDescription, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5200
                AddPCLproperty "Advisory", .Advisory, wdText, "Disabled"

'vbwLine 5201:            Case PERFORMANCEWHEEL
            Case IIf(vbwProfiler.vbwExecuteLine(5201), VBWPROFILER_EMPTY, _
        PERFORMANCEWHEEL)
vbwProfiler.vbwExecuteLine 5202
                AddPCLproperty "Wheeled Performance", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5203
                AddPCLproperty "gSpeed", Format(.gTopSpeed, "standard") & " mph", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5204
                AddPCLproperty "gOffRd", Format(.gOffRoad, "standard") & " mph", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5205
                AddPCLproperty "gAccel", .gAcceleration & " mph/s", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5206
                AddPCLproperty "gDecel", .gDeceleration & " mph/s", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5207
                AddPCLproperty "gSR", .gStability, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5208
                AddPCLproperty "gMR", .gManeuverability, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5209
                AddPCLproperty "gP", Format(.gPressure, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5210
                AddPCLproperty "gPDescr", .gPressureDescription, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5211
                AddPCLproperty "Advisory", .Advisory, wdText, "Disabled"

'vbwLine 5212:            Case PERFORMANCEFLEX
            Case IIf(vbwProfiler.vbwExecuteLine(5212), VBWPROFILER_EMPTY, _
        PERFORMANCEFLEX)
vbwProfiler.vbwExecuteLine 5213
                AddPCLproperty "Flexibody Performance", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5214
                AddPCLproperty "gSpeed", Format(.gTopSpeed, "standard") & " mph", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5215
                AddPCLproperty "gOffRd", Format(.gOffRoad, "standard") & " mph", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5216
                AddPCLproperty "gAccel", .gAcceleration & " mph/s", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5217
                AddPCLproperty "gDecel", .gDeceleration & " mph/s", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5218
                AddPCLproperty "gSR", .gStability, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5219
                AddPCLproperty "gMR", .gManeuverability, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5220
                AddPCLproperty "gP", Format(.gPressure, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5221
                AddPCLproperty "gPDescr", .gPressureDescription, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5222
                AddPCLproperty "Advisory", .Advisory, wdText, "Disabled"

'vbwLine 5223:            Case PERFORMANCESKID
            Case IIf(vbwProfiler.vbwExecuteLine(5223), VBWPROFILER_EMPTY, _
        PERFORMANCESKID)
vbwProfiler.vbwExecuteLine 5224
                AddPCLproperty "Skid Performance", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5225
                AddPCLproperty "gSpeed", Format(.gTopSpeed, "standard") & " mph", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5226
                AddPCLproperty "gOffRd", Format(.gOffRoad, "standard") & " mph", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5227
                AddPCLproperty "gAccel", .gAcceleration & " mph/s", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5228
                AddPCLproperty "gDecel", .gDeceleration & " mph/s", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5229
                AddPCLproperty "gSR", .gStability, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5230
                AddPCLproperty "gMR", .gManeuverability, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5231
                AddPCLproperty "gP", Format(.gPressure, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5232
                AddPCLproperty "gPDescr", .gPressureDescription, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5233
                AddPCLproperty "Advisory", .Advisory, wdText, "Disabled"

'vbwLine 5234:            Case PERFORMANCEAIR
            Case IIf(vbwProfiler.vbwExecuteLine(5234), VBWPROFILER_EMPTY, _
        PERFORMANCEAIR)
vbwProfiler.vbwExecuteLine 5235
                AddPCLproperty "Air Performance", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5236
                AddPCLproperty "Thrust Options", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5237
                AddPCLproperty "PercentThrust", .percentthrust, wdNumber, "PercentThrust"
vbwProfiler.vbwExecuteLine 5238
                AddPCLproperty "TreatTiltRotorsAsPropellers", .TreatTiltRotorsAsPropellers, wdBool, "TreatTiltRotorsAsPropellers"
vbwProfiler.vbwExecuteLine 5239
                AddPCLproperty "AfterBurnersOn", .AfterBurnersOn, wdBool, "AfterBurnersOn"


vbwProfiler.vbwExecuteLine 5240
                AddPCLproperty "Streamlining", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5241
                AddPCLproperty "HardPointsOn", .HardPointsOn, wdBool, "HardPointsOn"
vbwProfiler.vbwExecuteLine 5242
                AddPCLproperty "WheelsSkidsExtended", .WheelsSkidsExtended, wdBool, "WheelsSkidsExtended"
vbwProfiler.vbwExecuteLine 5243
                AddPCLproperty "PopTurretsExtended", .PopTurretsExtended, wdBool, "PopTurretsExtended"

vbwProfiler.vbwExecuteLine 5244
                AddPCLproperty "Weight Percentages", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5245
                AddPCLproperty "% Crew", .PercentCrewWeight, wdNumber, "PercentCrewWeight"
vbwProfiler.vbwExecuteLine 5246
                AddPCLproperty "% Fuel", .PercentFuelWeight, wdNumber, "PercentFuelWeight"
vbwProfiler.vbwExecuteLine 5247
                AddPCLproperty "% Cargo", .PercentCargoWeight, wdNumber, "PercentCargoWeight"
vbwProfiler.vbwExecuteLine 5248
                AddPCLproperty "% Hardpoints/Bays Load", .PercentHardpointWeight, wdNumber, "PercentHardpointWeight"
vbwProfiler.vbwExecuteLine 5249
                AddPCLproperty "% Provisions", .PercentProvisionWeight, wdNumber, "PercentProvisionWeight"
vbwProfiler.vbwExecuteLine 5250
                AddPCLproperty "% Ammunitions", .PercentAmmunitionWeight, wdNumber, "PercentAmmunitionWeight"
vbwProfiler.vbwExecuteLine 5251
                AddPCLproperty "% PercentAuxVehicleWeight", .PercentAuxVehicleWeight, wdNumber, "PercentAuxVehicleWeight"

vbwProfiler.vbwExecuteLine 5252
                AddPCLproperty "Statistics", "", wdText, PROPERTY_HEADER

                'frmDesigner.lblperformance(0).Caption =  "Can Fly?" & vbTab & .aCanFly
vbwProfiler.vbwExecuteLine 5253
                AddPCLproperty "Thrust", Format(.aMotiveThrust, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5254
                AddPCLproperty "Static Lift", Format(.staticlift, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5255
                AddPCLproperty "Drag", .aDrag, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5256
                AddPCLproperty "Speed", Format(.aTopSpeed, "standard") & " mph", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5257
                AddPCLproperty "Stall Speed", Format(.aStallSpeed, "standard") & " mph", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5258
                AddPCLproperty "aAccel", Format(.aAcceleration, "standard") & " mph/s", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5259
                AddPCLproperty "aDecel", Format(.aDeceleration, "standard") & " mph/s", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5260
                AddPCLproperty "aMR", .aManeuverability, wdText, "Disabled"

vbwProfiler.vbwExecuteLine 5261
                AddPCLproperty "aSR", .aStability, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5262
                AddPCLproperty "TakeOff Run (yrds)", .aTakeOffRun, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5263
                AddPCLproperty "Landing Run (yrds)", .aLandingRun, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5264
                AddPCLproperty "Advisory", .Advisory, wdText, "Disabled"


'vbwLine 5265:            Case PERFORMANCEHOVER
            Case IIf(vbwProfiler.vbwExecuteLine(5265), VBWPROFILER_EMPTY, _
        PERFORMANCEHOVER)
vbwProfiler.vbwExecuteLine 5266
                AddPCLproperty "Hovercraft Performance", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5267
                AddPCLproperty "Hover Alt", .hHoverAltitude & " feet", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5268
                AddPCLproperty "Thrust", Format(.hMotiveThrust, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5269
                AddPCLproperty "Static Lift", Format(.staticlift, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5270
                AddPCLproperty "Speed", Format(.hTopSpeed, "standard") & " mph", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5271
                AddPCLproperty "Drag", .hDrag, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5272
                AddPCLproperty "hAccel", Format(.hAcceleration, "standard") & " mph/s", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5273
                AddPCLproperty "hDecel", Format(.hDeceleration, "standard") & " mph/s", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5274
                AddPCLproperty "hSR", .hstability, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5275
                AddPCLproperty "hMR", .hmaneuverability & " g", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5276
                AddPCLproperty "Advisory", .Advisory, wdText, "Disabled"


'vbwLine 5277:            Case PERFORMANCEMAGLEV
            Case IIf(vbwProfiler.vbwExecuteLine(5277), VBWPROFILER_EMPTY, _
        PERFORMANCEMAGLEV)
vbwProfiler.vbwExecuteLine 5278
                AddPCLproperty "Mag-Lev Performance", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5279
                AddPCLproperty "mThrust", Format(.mlMotiveThrust, "standard") & " lbs", wdText, "Disabled"
                'todo: StaticLift?  'AddPCLproperty "Static Lift", Format(.mlstaticlift, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5280
                AddPCLproperty "mSpeed", Format(.mlTopSpeed, "standard") & " mph", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5281
                AddPCLproperty "Stall Speed", Format(.mlStallSpeed, "standard") & " mph", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5282
                AddPCLproperty "mDrag", .mlDrag, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5283
                AddPCLproperty "mAccel", Format(.mlAcceleration, "standard") & " mph/s", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5284
                AddPCLproperty "mDecel", Format(.mlDeceleration, "standard") & " mph/s", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5285
                AddPCLproperty "mSR", .mlStability, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5286
                AddPCLproperty "mMR", .mlManeuverability, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5287
                AddPCLproperty "Advisory", .Advisory, wdText, "Disabled"


'vbwLine 5288:            Case PERFORMANCEWATER
            Case IIf(vbwProfiler.vbwExecuteLine(5288), VBWPROFILER_EMPTY, _
        PERFORMANCEWATER)
vbwProfiler.vbwExecuteLine 5289
                AddPCLproperty "Water Performance", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5290
                AddPCLproperty "wThrust", Format(.wTotalAquaticThrust, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5291
                AddPCLproperty "wDrag", Format(.wHydroDrag, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5292
                AddPCLproperty "wSpeed", Format(.wTopSpeed, "standard") & " mph", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5293
                AddPCLproperty "Hydro Speed", Format(.wHydrofoilSpeed, "standard") & " mph", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5294
                AddPCLproperty "Planing Speed", Format(.wPlaningSpeed, "standard") & " mph", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5295
                AddPCLproperty "wAccel", Format(.wAcceleration, "standard") & " mph/s", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5296
                AddPCLproperty "wDecel", Format(.wDeceleration, "standard") & " mph/s", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5297
                AddPCLproperty "Incr Decel", Format(.wIDeceleration, "standard") & " mph/s", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5298
                AddPCLproperty "wSR", .wStability & "  " & "wMR: " & .wManeuverability, wdText, "Disabled"
                'AddPCLproperty  "wMR",  .wManeuverability, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5299
                AddPCLproperty "wDraft", Format(.wDraft, "standard") & " feet", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5300
                AddPCLproperty "Advisory", .Advisory, wdText, "Disabled"


'vbwLine 5301:            Case PERFORMANCESUB
            Case IIf(vbwProfiler.vbwExecuteLine(5301), VBWPROFILER_EMPTY, _
        PERFORMANCESUB)
vbwProfiler.vbwExecuteLine 5302
                AddPCLproperty "Submerged Performance", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5303
                AddPCLproperty "suThrust", Format(.suTotalAquaticThrust, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5304
                AddPCLproperty "suDrag", .suHydroDrag, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5305
                AddPCLproperty "suSpeed", Format(.suTopSpeed, "standard") & " mph", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5306
                AddPCLproperty "suAccel", Format(.suAcceleration, "standard") & " mph/s", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5307
                AddPCLproperty "suDecel", Format(.suDeceleration, "standard") & " mph/s", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5308
                AddPCLproperty "Incr Decel", Format(.suIDeceleration, "standard") & " mph/s", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5309
                AddPCLproperty "suSR", .suStability, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5310
                AddPCLproperty "suMR", .suManeuverability, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5311
                AddPCLproperty "Draft", Format(.suDraft, "standard") & " feet", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5312
                If .suCrushDepth = -1 Then
vbwProfiler.vbwExecuteLine 5313
                    AddPCLproperty "Crush Depth", "No Crush Depth", wdText, "Disabled"
                Else
vbwProfiler.vbwExecuteLine 5314 'B
vbwProfiler.vbwExecuteLine 5315
                    AddPCLproperty "Crush Depth", Format(.suCrushDepth, "standard") & " yards", wdText, "Disabled"
                End If
vbwProfiler.vbwExecuteLine 5316 'B
vbwProfiler.vbwExecuteLine 5317
                AddPCLproperty "Advisory", .Advisory, wdText, "Disabled"


'vbwLine 5318:           Case PERFORMANCESPACE
           Case IIf(vbwProfiler.vbwExecuteLine(5318), VBWPROFILER_EMPTY, _
        PERFORMANCESPACE)
vbwProfiler.vbwExecuteLine 5319
                AddPCLproperty "Space Performance", "", wdText, PROPERTY_HEADER
vbwProfiler.vbwExecuteLine 5320
                AddPCLproperty "Thrust", Format(.sMotiveThrust, "standard") & " lbs", wdText, "Disabled"
                'todo: make sure accel is displaying at least 4 digits for space craft
                'which have very slow accel but eventually build up to very fast speeds.
vbwProfiler.vbwExecuteLine 5321
                AddPCLproperty "sAccel", Format(.sAccelerationG, "###,###,###.####") & " g", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5322
                AddPCLproperty "sAccel", Format(.sAccelerationMPH, "standard") & " mph/s", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5323
                AddPCLproperty "Turn Around", Format(.sTurnAroundTime, "standard") & " secs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5324
                AddPCLproperty "sMR", Format(.sManeuverability, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5325
                AddPCLproperty "Hyper", Format(.sHyperSpeed, "standard") & " parsecs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5326
                AddPCLproperty "Warp", Format(.sWarpSpeed, "standard") & " parsecs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5327
                AddPCLproperty "Jump?", .sJumpDriveable, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5328
                AddPCLproperty "Teleport?", .sTeleportationDriveable, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5329
                AddPCLproperty "Advisory", .Advisory, wdText, "Disabled"
        End Select
vbwProfiler.vbwExecuteLine 5330 'B
vbwProfiler.vbwExecuteLine 5331
End With

vbwProfiler.vbwProcOut 268
vbwProfiler.vbwExecuteLine 5332
End Sub

Private Sub ShowPropsForGroupComponent(ByVal component As Integer, ByVal Key As String)
vbwProfiler.vbwProcIn 269

vbwProfiler.vbwExecuteLine 5333
    With m_oCurrentVeh.Components(Key)
vbwProfiler.vbwExecuteLine 5334
        AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5335
    End With

vbwProfiler.vbwProcOut 269
vbwProfiler.vbwExecuteLine 5336
End Sub

Private Sub ShowPropsForSimpleCustom(ByVal component As Integer, ByVal Key As String)
vbwProfiler.vbwProcIn 270
vbwProfiler.vbwExecuteLine 5337
    Select Case component

'vbwLine 5338:        Case SimpleCustom
        Case IIf(vbwProfiler.vbwExecuteLine(5338), VBWPROFILER_EMPTY, _
        SimpleCustom)
vbwProfiler.vbwExecuteLine 5339
            With m_oCurrentVeh.Components(Key)
vbwProfiler.vbwExecuteLine 5340
                AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5341
                AddPCLproperty "Tech level", .TL, wdList, "TL", 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 5342
                AddPCLproperty "User Cost", .UserCost, wdDouble, "UserCost"
vbwProfiler.vbwExecuteLine 5343
                AddPCLproperty "User Weight", .UserWeight, wdDouble, "UserWeight"
vbwProfiler.vbwExecuteLine 5344
                AddPCLproperty "User Volume", .UserVolume, wdDouble, "UserVolume"
vbwProfiler.vbwExecuteLine 5345
                AddPCLproperty "Power Consumption", .PowerReqt, wdDouble, "PowerReqt"
vbwProfiler.vbwExecuteLine 5346
                AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 5347
                AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 5348
                AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5349
                AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5350
                AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5351
                AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5352
                AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5353
                AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

vbwProfiler.vbwExecuteLine 5354
            End With
    End Select
vbwProfiler.vbwExecuteLine 5355 'B

vbwProfiler.vbwProcOut 270
vbwProfiler.vbwExecuteLine 5356
End Sub

Public Sub ShowPropsForWeaponLink(ByRef sKey As String)
vbwProfiler.vbwProcIn 271


vbwProfiler.vbwExecuteLine 5357
    With m_oCurrentVeh.WeaponProfiles(sKey)
vbwProfiler.vbwExecuteLine 5358
        AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5359
    End With
vbwProfiler.vbwProcOut 271
vbwProfiler.vbwExecuteLine 5360
End Sub
Private Sub ShowPropsForWeaponry1(ByVal component As Integer, ByVal Key As String)
vbwProfiler.vbwProcIn 272
Dim listarray() As String

vbwProfiler.vbwExecuteLine 5361
With m_oCurrentVeh.Components(Key)

' Fill the window with properties for the correct Collection item
vbwProfiler.vbwExecuteLine 5362
Select Case component

'vbwLine 5363:    Case BlueGreenLaser, RainbowLaser, Laser, UVLaser, IRLaser, Disruptor, ChargedParticleBeam, NeutralParticleBeam, Flamer, Screamer, Stunner, ParalysisBeam, XRayLaser, FusionBeam, GravityBeam, AntiparticleBeam, Graser, Disintegrator, Displacer, BeamedPowerTransmitter, MilitaryParalysisBeam
    Case IIf(vbwProfiler.vbwExecuteLine(5363), VBWPROFILER_EMPTY, _
        BlueGreenLaser), RainbowLaser, Laser, UVLaser, IRLaser, Disruptor, ChargedParticleBeam, NeutralParticleBeam, Flamer, Screamer, Stunner, ParalysisBeam, XRayLaser, FusionBeam, GravityBeam, AntiparticleBeam, Graser, Disintegrator, Displacer, BeamedPowerTransmitter, MilitaryParalysisBeam
vbwProfiler.vbwExecuteLine 5364
       AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5365
       AddPCLproperty "Tech level", .TL, wdList, "TL", 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 5366
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 5367
       AddPCLproperty "Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)"
vbwProfiler.vbwExecuteLine 5368
       AddPCLproperty "Mounting", .Mount, wdList, "Mount", "normal", "concealed"
vbwProfiler.vbwExecuteLine 5369
       AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
vbwProfiler.vbwExecuteLine 5370
       If (component = BlueGreenLaser) Or (component = RainbowLaser) Or (component = Disintegrator) Or (component = Flamer) Or (component = Laser) Then
vbwProfiler.vbwExecuteLine 5371
            AddPCLproperty "Energy Drill", .EnergyDrill, wdBool, "EnergyDrill"
       End If
vbwProfiler.vbwExecuteLine 5372 'B
vbwProfiler.vbwExecuteLine 5373
       AddPCLproperty "Beam Output", .BeamOutput, wdDouble, "BeamOutput"
vbwProfiler.vbwExecuteLine 5374
       AddPCLproperty "Cyclic Rate", .rof, wdList, "rof", "1/14400", "1/7200", "1/4800", "1/3600", "1/2400", "1/1200", "1/600", "1/300", "1/150", "1/60", "1/30", "1/15", "1/8", "1/4", "1/2", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20"
vbwProfiler.vbwExecuteLine 5375
       AddPCLproperty "Range", .Range, wdList, "Range", "close", "normal", "long", "very long", "extreme"
vbwProfiler.vbwExecuteLine 5376
       AddPCLproperty "Power Cells", .PowerCellType, wdList, "PowerCellType", "none", "C cells", "rC cell", "D cells", "rD cells", "E cells", "rE cells"
vbwProfiler.vbwExecuteLine 5377
       AddPCLproperty "# Power Cells", .PowerCellQuantity, wdNumber, "PowerCellQuantity"
vbwProfiler.vbwExecuteLine 5378
       AddPCLproperty "FTL Beam?", .FTL, wdBool, "FTL"
vbwProfiler.vbwExecuteLine 5379
       AddPCLproperty "Compact", .Compact, wdBool, "Compact"
vbwProfiler.vbwExecuteLine 5380
       AddPCLproperty "Reputation for Quality", .Reliable, wdBool, "Reliable"
vbwProfiler.vbwExecuteLine 5381
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 5382
       AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 5383
       AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5384
       AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5385
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5386
       AddPCLproperty "Type Damage", .TypeDamage, wdText, "Disabled"
        'note: for displacers its radius of effect and not damage
        'note: for paralysis beams its HT penalty and not damage
        'note: for stunners its HT penalty also and not damage
vbwProfiler.vbwExecuteLine 5387
       If (component = Stunner) Or (component = ParalysisBeam) Or (component = MilitaryParalysisBeam) Then
vbwProfiler.vbwExecuteLine 5388
            AddPCLproperty "HT penalty", .Damage, wdText, "Disabled"
       Else
vbwProfiler.vbwExecuteLine 5389 'B
vbwProfiler.vbwExecuteLine 5390
           AddPCLproperty "Damage", .Damage, wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 5391 'B
        'note: militaryparalysis,paralysis, disintegrators and displacers have no half damage
vbwProfiler.vbwExecuteLine 5392
       AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5393
       AddPCLproperty "Vacuum 1/2 Damage (yards)", .VacuumHalfDamage, wdText, "Disabled"

vbwProfiler.vbwExecuteLine 5394
        If component = Stunner Then
vbwProfiler.vbwExecuteLine 5395
           AddPCLproperty "Max Range at HT6- (yards)", .MaxRange, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5396
           AddPCLproperty "Max Range at HT7+ (yards)", .MaxRange2, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5397
           AddPCLproperty "Vacuum Max Range at HT6- (yards)", .VacuumMaxRange, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5398
           AddPCLproperty "Vacuum Max Range at HT7+ (yards)", .VacuumMaxRange2, wdText, "Disabled"
        Else
vbwProfiler.vbwExecuteLine 5399 'B
vbwProfiler.vbwExecuteLine 5400
           AddPCLproperty "Max Range (yards)", .MaxRange, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5401
           AddPCLproperty "Vacuum Max Range (yards)", .VacuumMaxRange, wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 5402 'B
vbwProfiler.vbwExecuteLine 5403
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5404
       AddPCLproperty "Snap Shot", .SnapShot, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5405
       AddPCLproperty "Rate of Fire", .rof, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5406
       AddPCLproperty "# of Shots", .Shots, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5407
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5408
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5409
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5410
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5411
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 5412:     Case UnGuidedMissile, UnGuidedTorpedo
     Case IIf(vbwProfiler.vbwExecuteLine(5412), VBWPROFILER_EMPTY, _
        UnGuidedMissile), UnGuidedTorpedo
vbwProfiler.vbwExecuteLine 5413
       AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5414
       AddPCLproperty "Tech level", .TL, wdList, "TL", 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 5415
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 5416
       AddPCLproperty "Diameter", .Diameter, wdDouble, "Diameter"
vbwProfiler.vbwExecuteLine 5417
       listarray = .FillAmmunitionList
vbwProfiler.vbwExecuteLine 5418
       AddPCLproperty "Warhead", .WarHead, wdList, "Warhead", listarray
vbwProfiler.vbwExecuteLine 5419
       AddPCLproperty "Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge"
vbwProfiler.vbwExecuteLine 5420
       AddPCLproperty "# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"
vbwProfiler.vbwExecuteLine 5421
        If .SpaceMissile = False Then
vbwProfiler.vbwExecuteLine 5422
           AddPCLproperty "Speed (yds per sec)", .Speed, wdDouble, "Speed"
        Else
vbwProfiler.vbwExecuteLine 5423 'B
vbwProfiler.vbwExecuteLine 5424
           AddPCLproperty "G's", .Speed, wdDouble, "Speed"
        End If
vbwProfiler.vbwExecuteLine 5425 'B
vbwProfiler.vbwExecuteLine 5426
       AddPCLproperty "Motor Weight", .MotorWeight, wdDouble, "MotorWeight"
vbwProfiler.vbwExecuteLine 5427
       AddPCLproperty "Stealth", .Stealth, wdBool, "Stealth"
        'this option only available for relevant missile types
vbwProfiler.vbwExecuteLine 5428
       AddPCLproperty "Space Fairing?", .SpaceMissile, wdBool, "SpaceMissile"
vbwProfiler.vbwExecuteLine 5429
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 5430
       AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 5431
       AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5432
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5433
       If .BurstRadius <> -1 Then
vbwProfiler.vbwExecuteLine 5434
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
vbwProfiler.vbwExecuteLine 5435 'B
vbwProfiler.vbwExecuteLine 5436
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5437
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5438
        If .TypeDamage2 <> "none" Then
vbwProfiler.vbwExecuteLine 5439
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5440
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 5441 'B
vbwProfiler.vbwExecuteLine 5442
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5443
       AddPCLproperty "Endurance (seconds)", .Endurance, wdText, "Disabled"
        'only unguided missiles have 1/2 damage
vbwProfiler.vbwExecuteLine 5444
        If component = UnGuidedMissile Then
vbwProfiler.vbwExecuteLine 5445
           AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 5446 'B
vbwProfiler.vbwExecuteLine 5447
       AddPCLproperty "Max Range (yards)", .MaxRange, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5448
       AddPCLproperty "Min Range (yards)", .MinRange, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5449
       AddPCLproperty "Motor Cost", "$" & Format(.MotorCost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5450
       AddPCLproperty "Warhead Cost", "$" & Format(.WarheadCost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5451
       AddPCLproperty "Warhead Weight", Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5452
       AddPCLproperty "Payload Cost", "$" & Format(.PayloadCost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5453
       AddPCLproperty "Payload Weight", Format(.PayloadWeight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5454
       AddPCLproperty "Total Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5455
       AddPCLproperty "Total Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5456
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5457
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5458
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 5459:    Case GuidedMissile, GuidedTorpedo
    Case IIf(vbwProfiler.vbwExecuteLine(5459), VBWPROFILER_EMPTY, _
        GuidedMissile), GuidedTorpedo
vbwProfiler.vbwExecuteLine 5460
       AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5461
       AddPCLproperty "Tech level", .TL, wdList, "TL", 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 5462
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 5463
       listarray = .FillGuidanceList
vbwProfiler.vbwExecuteLine 5464
       AddPCLproperty "Guidance System", .GuidanceSystem, wdList, "GuidanceSystem", listarray
vbwProfiler.vbwExecuteLine 5465
       listarray = .FillTerminalGuidanceList
vbwProfiler.vbwExecuteLine 5466
       AddPCLproperty "Terminal Guidance", .BrilliantGuidanceSystem, wdList, "BrilliantGuidanceSystem", listarray
vbwProfiler.vbwExecuteLine 5467
       AddPCLproperty "Cheap Guidance System", .CheapGuidance, wdBool, "CheapGuidance"
vbwProfiler.vbwExecuteLine 5468
       AddPCLproperty "Compact Guidance System", .Compact, wdBool, "Compact"
vbwProfiler.vbwExecuteLine 5469
       AddPCLproperty "Mid-Course Update", .MidCourseUpdate, wdBool, "MidCourseUpdate"
vbwProfiler.vbwExecuteLine 5470
       AddPCLproperty "Pop-Up", .PopUp, wdBool, "Popup"
vbwProfiler.vbwExecuteLine 5471
       AddPCLproperty "Diameter", .Diameter, wdDouble, "Diameter"
vbwProfiler.vbwExecuteLine 5472
       listarray = .FillAmmunitionList
vbwProfiler.vbwExecuteLine 5473
       AddPCLproperty "Warhead", .WarHead, wdList, "Warhead", listarray
vbwProfiler.vbwExecuteLine 5474
       AddPCLproperty "Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge"
vbwProfiler.vbwExecuteLine 5475
       AddPCLproperty "# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"
vbwProfiler.vbwExecuteLine 5476
       AddPCLproperty "Skill Bonus", .SkillBonus, wdNumber, "SkillBonus"
        'this option only available for relevant missile /torp types
vbwProfiler.vbwExecuteLine 5477
        If .SpaceMissile = False Then
vbwProfiler.vbwExecuteLine 5478
           AddPCLproperty "Speed (yds per sec)", .Speed, wdDouble, "Speed"
        Else
vbwProfiler.vbwExecuteLine 5479 'B
vbwProfiler.vbwExecuteLine 5480
           AddPCLproperty "G's", .Speed, wdDouble, "Speed"
        End If
vbwProfiler.vbwExecuteLine 5481 'B
vbwProfiler.vbwExecuteLine 5482
       AddPCLproperty "Motor Weight", .MotorWeight, wdDouble, "MotorWeight"
vbwProfiler.vbwExecuteLine 5483
       AddPCLproperty "Stealth", .Stealth, wdBool, "Stealth"
        'this option only available for relevant missile types
vbwProfiler.vbwExecuteLine 5484
       AddPCLproperty "Space Fairing?", .SpaceMissile, wdBool, "SpaceMissile"
vbwProfiler.vbwExecuteLine 5485
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 5486
       AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 5487
       AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5488
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5489
       If .BurstRadius <> -1 Then
vbwProfiler.vbwExecuteLine 5490
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
vbwProfiler.vbwExecuteLine 5491 'B
vbwProfiler.vbwExecuteLine 5492
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5493
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5494
        If .TypeDamage2 <> "none" Then
vbwProfiler.vbwExecuteLine 5495
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5496
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 5497 'B
vbwProfiler.vbwExecuteLine 5498
       AddPCLproperty "Skill", .Skill, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5499
       AddPCLproperty "Endurance (seconds)", .Endurance, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5500
       AddPCLproperty "Max Range (yards)", .MaxRange, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5501
       AddPCLproperty "Min Range (yards)", .MinRange, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5502
       AddPCLproperty "Motor Cost", "$" & Format(.MotorCost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5503
       AddPCLproperty "Warhead Cost", "$" & Format(.WarheadCost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5504
       AddPCLproperty "Warhead Weight", Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5505
       AddPCLproperty "Guidance System Cost", "$" & Format(.GuidanceCost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5506
       AddPCLproperty "Guidance System Weight", Format(.GuidanceWeight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5507
       AddPCLproperty "Payload Cost", "$" & Format(.PayloadCost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5508
       AddPCLproperty "Payload Weight", Format(.PayloadWeight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5509
       AddPCLproperty "Total Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5510
       AddPCLproperty "Total Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5511
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5512
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5513
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

'vbwLine 5514:    Case IronBomb, SelfDestructSystem
    Case IIf(vbwProfiler.vbwExecuteLine(5514), VBWPROFILER_EMPTY, _
        IronBomb), SelfDestructSystem
vbwProfiler.vbwExecuteLine 5515
       AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5516
       AddPCLproperty "Tech level", .TL, wdList, "TL", 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 5517
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 5518
       AddPCLproperty "Diameter", .Diameter, wdDouble, "Diameter"
vbwProfiler.vbwExecuteLine 5519
       listarray = .FillAmmunitionList
vbwProfiler.vbwExecuteLine 5520
       AddPCLproperty "Warhead", .WarHead, wdList, "Warhead", listarray
vbwProfiler.vbwExecuteLine 5521
       AddPCLproperty "Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge"
vbwProfiler.vbwExecuteLine 5522
       AddPCLproperty "# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"
       'note: these dont have a speed do they?
       'AddPCLproperty "Speed (yds per sec)", .Speed, wdDouble, "Speed"
vbwProfiler.vbwExecuteLine 5523
       AddPCLproperty "Stealth", .Stealth, wdBool, "Stealth"
vbwProfiler.vbwExecuteLine 5524
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 5525
       AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 5526
       AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5527
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5528
       If .BurstRadius <> -1 Then
vbwProfiler.vbwExecuteLine 5529
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
vbwProfiler.vbwExecuteLine 5530 'B
vbwProfiler.vbwExecuteLine 5531
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5532
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5533
        If .TypeDamage2 <> "none" Then
vbwProfiler.vbwExecuteLine 5534
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5535
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 5536 'B
vbwProfiler.vbwExecuteLine 5537
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5538
       AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5539
       AddPCLproperty "Min Range (yards)", .MinRange, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5540
       AddPCLproperty "Warhead Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5541
       AddPCLproperty "Warhead Weight", Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5542
       AddPCLproperty "Cost", "$" & Format(.PayloadCost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5543
       AddPCLproperty "Weight", Format(.PayloadWeight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5544
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5545
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5546
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 5547:    Case RetardedBomb
    Case IIf(vbwProfiler.vbwExecuteLine(5547), VBWPROFILER_EMPTY, _
        RetardedBomb)
vbwProfiler.vbwExecuteLine 5548
       AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5549
       AddPCLproperty "Tech level", .TL, wdList, "TL", 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 5550
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 5551
       AddPCLproperty "Diameter", .Diameter, wdDouble, "Diameter"
vbwProfiler.vbwExecuteLine 5552
       listarray = .FillAmmunitionList
vbwProfiler.vbwExecuteLine 5553
       AddPCLproperty "Warhead", .WarHead, wdList, "Warhead", listarray
vbwProfiler.vbwExecuteLine 5554
       AddPCLproperty "Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge"
vbwProfiler.vbwExecuteLine 5555
       AddPCLproperty "# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"
vbwProfiler.vbwExecuteLine 5556
       AddPCLproperty "Speed (yds per sec)", .Speed, wdDouble, "Speed"
vbwProfiler.vbwExecuteLine 5557
       AddPCLproperty "Stealth", .Stealth, wdBool, "Stealth"
vbwProfiler.vbwExecuteLine 5558
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 5559
       AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 5560
       AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5561
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5562
       If .BurstRadius <> -1 Then
vbwProfiler.vbwExecuteLine 5563
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
vbwProfiler.vbwExecuteLine 5564 'B
vbwProfiler.vbwExecuteLine 5565
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5566
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5567
        If .TypeDamage2 <> "none" Then
vbwProfiler.vbwExecuteLine 5568
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5569
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 5570 'B
vbwProfiler.vbwExecuteLine 5571
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5572
       AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5573
       AddPCLproperty "Min Range (yards)", .MinRange, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5574
       AddPCLproperty "Warhead Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5575
       AddPCLproperty "Payload Cost", "$" & Format(.PayloadCost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5576
       AddPCLproperty "Warhead Weight", Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5577
       AddPCLproperty "Payload Weight", Format(.PayloadWeight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5578
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5579
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5580
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 5581:    Case SmartBomb
    Case IIf(vbwProfiler.vbwExecuteLine(5581), VBWPROFILER_EMPTY, _
        SmartBomb)
vbwProfiler.vbwExecuteLine 5582
       AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5583
       AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 5584
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 5585
       listarray = .FillGuidanceList
vbwProfiler.vbwExecuteLine 5586
       AddPCLproperty "Guidance System", .GuidanceSystem, wdList, "GuidanceSystem", listarray
vbwProfiler.vbwExecuteLine 5587
       AddPCLproperty "Cheap Guidance System", .CheapGuidance, wdBool, "CheapGuidance"
vbwProfiler.vbwExecuteLine 5588
       AddPCLproperty "Compact Guidance System", .Compact, wdBool, "Compact"
vbwProfiler.vbwExecuteLine 5589
       AddPCLproperty "Diameter", .Diameter, wdDouble, "Diameter"
vbwProfiler.vbwExecuteLine 5590
       listarray = .FillAmmunitionList
vbwProfiler.vbwExecuteLine 5591
       AddPCLproperty "Warhead", .WarHead, wdList, "Warhead", listarray
vbwProfiler.vbwExecuteLine 5592
       AddPCLproperty "Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge"
vbwProfiler.vbwExecuteLine 5593
       AddPCLproperty "# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"
vbwProfiler.vbwExecuteLine 5594
       AddPCLproperty "Skill Bonus", .SkillBonus, wdNumber, "SkillBonus"
vbwProfiler.vbwExecuteLine 5595
       AddPCLproperty "Speed (yds per sec)", .Speed, wdDouble, "Speed"
vbwProfiler.vbwExecuteLine 5596
       AddPCLproperty "Stealth", .Stealth, wdBool, "Stealth"
vbwProfiler.vbwExecuteLine 5597
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 5598
       AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 5599
       AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5600
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5601
       If .BurstRadius <> -1 Then
vbwProfiler.vbwExecuteLine 5602
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
vbwProfiler.vbwExecuteLine 5603 'B
vbwProfiler.vbwExecuteLine 5604
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5605
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5606
        If .TypeDamage2 <> "none" Then
vbwProfiler.vbwExecuteLine 5607
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5608
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 5609 'B
vbwProfiler.vbwExecuteLine 5610
       AddPCLproperty "Skill", .Skill, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5611
       AddPCLproperty "Min Range (yards)", .MinRange, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5612
       AddPCLproperty "Warhead Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5613
       AddPCLproperty "Guidance System Cost", "$" & Format(.GuidanceCost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5614
       AddPCLproperty "Payload Cost", "$" & Format(.PayloadCost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5615
       AddPCLproperty "Warhead Weight", Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5616
       AddPCLproperty "Guidance System Weight", Format(.GuidanceWeight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5617
       AddPCLproperty "Payload Weight", Format(.PayloadWeight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5618
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5619
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5620
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 5621:    Case ContactMine
    Case IIf(vbwProfiler.vbwExecuteLine(5621), VBWPROFILER_EMPTY, _
        ContactMine)
vbwProfiler.vbwExecuteLine 5622
       AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5623
       AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 5624
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 5625
       AddPCLproperty "Diameter", .Diameter, wdDouble, "Diameter"
vbwProfiler.vbwExecuteLine 5626
       listarray = .FillAmmunitionList
vbwProfiler.vbwExecuteLine 5627
       AddPCLproperty "Warhead", .WarHead, wdList, "Warhead", listarray
vbwProfiler.vbwExecuteLine 5628
       AddPCLproperty "Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge"
vbwProfiler.vbwExecuteLine 5629
       AddPCLproperty "# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"
vbwProfiler.vbwExecuteLine 5630
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 5631
       AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 5632
       AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5633
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5634
       If .BurstRadius <> -1 Then
vbwProfiler.vbwExecuteLine 5635
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
vbwProfiler.vbwExecuteLine 5636 'B
vbwProfiler.vbwExecuteLine 5637
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5638
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5639
        If .TypeDamage2 <> "none" Then
vbwProfiler.vbwExecuteLine 5640
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5641
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 5642 'B
vbwProfiler.vbwExecuteLine 5643
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5644
       AddPCLproperty "Warhead Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5645
       AddPCLproperty "Warhead Weight", Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5646
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5647
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5648
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

'vbwLine 5649:    Case ProximityMine
    Case IIf(vbwProfiler.vbwExecuteLine(5649), VBWPROFILER_EMPTY, _
        ProximityMine)
vbwProfiler.vbwExecuteLine 5650
       AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5651
       AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 5652
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 5653
       listarray = .FillGuidanceList
vbwProfiler.vbwExecuteLine 5654
       AddPCLproperty "Guidance System", .GuidanceSystem, wdList, "GuidanceSystem", listarray
vbwProfiler.vbwExecuteLine 5655
       AddPCLproperty "Cheap Guidance System", .CheapGuidance, wdBool, "CheapGuidance"
vbwProfiler.vbwExecuteLine 5656
       AddPCLproperty "Compact Guidance System", .Compact, wdBool, "Compact"
vbwProfiler.vbwExecuteLine 5657
       AddPCLproperty "Diameter", .Diameter, wdDouble, "Diameter"
vbwProfiler.vbwExecuteLine 5658
       listarray = .FillAmmunitionList
vbwProfiler.vbwExecuteLine 5659
       AddPCLproperty "Warhead", .WarHead, wdList, "Warhead", listarray
vbwProfiler.vbwExecuteLine 5660
       AddPCLproperty "Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge"
vbwProfiler.vbwExecuteLine 5661
       AddPCLproperty "# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"
vbwProfiler.vbwExecuteLine 5662
       AddPCLproperty "Skill Bonus", .SkillBonus, wdNumber, "SkillBonus"
vbwProfiler.vbwExecuteLine 5663
       AddPCLproperty "Stealth", .Stealth, wdBool, "Stealth"
vbwProfiler.vbwExecuteLine 5664
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 5665
       AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 5666
       AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5667
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5668
       If .BurstRadius <> -1 Then
vbwProfiler.vbwExecuteLine 5669
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
vbwProfiler.vbwExecuteLine 5670 'B
vbwProfiler.vbwExecuteLine 5671
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5672
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5673
        If .TypeDamage2 <> "none" Then
vbwProfiler.vbwExecuteLine 5674
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5675
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 5676 'B
vbwProfiler.vbwExecuteLine 5677
       AddPCLproperty "Skill", .Skill, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5678
       AddPCLproperty "Warhead Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5679
       AddPCLproperty "Guidance System Cost", "$" & Format(.GuidanceCost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5680
       AddPCLproperty "Payload Cost", "$" & Format(.PayloadCost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5681
       AddPCLproperty "Warhead Weight", Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5682
       AddPCLproperty "Guidance System Weight", Format(.GuidanceWeight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5683
       AddPCLproperty "Payload Weight", Format(.PayloadWeight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5684
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5685
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5686
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 5687:    Case PressureTriggerMine
    Case IIf(vbwProfiler.vbwExecuteLine(5687), VBWPROFILER_EMPTY, _
        PressureTriggerMine)
vbwProfiler.vbwExecuteLine 5688
       AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5689
       AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 5690
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 5691
       AddPCLproperty "Diameter", .Diameter, wdDouble, "Diameter"
vbwProfiler.vbwExecuteLine 5692
       AddPCLproperty "Detonation Weight", .DetonationWeight, wdDouble, "DetonationWeight"
vbwProfiler.vbwExecuteLine 5693
       listarray = .FillAmmunitionList
vbwProfiler.vbwExecuteLine 5694
       AddPCLproperty "Warhead", .WarHead, wdList, "Warhead", listarray
vbwProfiler.vbwExecuteLine 5695
       AddPCLproperty "Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge"
vbwProfiler.vbwExecuteLine 5696
       AddPCLproperty "# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"
vbwProfiler.vbwExecuteLine 5697
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 5698
       AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 5699
       AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5700
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5701
       If .BurstRadius <> -1 Then
vbwProfiler.vbwExecuteLine 5702
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
vbwProfiler.vbwExecuteLine 5703 'B
vbwProfiler.vbwExecuteLine 5704
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5705
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5706
        If .TypeDamage2 <> "none" Then
vbwProfiler.vbwExecuteLine 5707
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5708
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 5709 'B
vbwProfiler.vbwExecuteLine 5710
       AddPCLproperty "Warhead Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5711
       AddPCLproperty "Warhead Weight", Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5712
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5713
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5714
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 5715:    Case CommandTriggerMine, SmartTriggerMine
    Case IIf(vbwProfiler.vbwExecuteLine(5715), VBWPROFILER_EMPTY, _
        CommandTriggerMine), SmartTriggerMine
vbwProfiler.vbwExecuteLine 5716
       AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5717
       AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 5718
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 5719
       AddPCLproperty "Diameter", .Diameter, wdDouble, "Diameter"
vbwProfiler.vbwExecuteLine 5720
       AddPCLproperty "Parachute Mine?", .Parachute, wdBool, "Parachute"
vbwProfiler.vbwExecuteLine 5721
       listarray = .FillAmmunitionList
vbwProfiler.vbwExecuteLine 5722
       AddPCLproperty "Warhead", .WarHead, wdList, "Warhead", listarray
vbwProfiler.vbwExecuteLine 5723
       AddPCLproperty "Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge"
vbwProfiler.vbwExecuteLine 5724
       AddPCLproperty "# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"
vbwProfiler.vbwExecuteLine 5725
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 5726
       AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 5727
       AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5728
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5729
       If .BurstRadius <> -1 Then
vbwProfiler.vbwExecuteLine 5730
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
vbwProfiler.vbwExecuteLine 5731 'B
vbwProfiler.vbwExecuteLine 5732
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5733
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5734
        If .TypeDamage2 <> "none" Then
vbwProfiler.vbwExecuteLine 5735
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5736
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 5737 'B
vbwProfiler.vbwExecuteLine 5738
       AddPCLproperty "Warhead Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5739
       AddPCLproperty "Warhead Weight", Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5740
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5741
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5742
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 5743:    Case WaterCannon, FlameThrower
    Case IIf(vbwProfiler.vbwExecuteLine(5743), VBWPROFILER_EMPTY, _
        WaterCannon), FlameThrower
vbwProfiler.vbwExecuteLine 5744
       AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5745
       AddPCLproperty "Tech level", .TL, wdList, "TL", 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 5746
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 5747
       AddPCLproperty "Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)"
vbwProfiler.vbwExecuteLine 5748
       AddPCLproperty "Mounting", .Mount, wdList, "Mount", "normal", "concealed"
vbwProfiler.vbwExecuteLine 5749
       AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
vbwProfiler.vbwExecuteLine 5750
       AddPCLproperty "Style", .Style, wdList, "Style", "light", "medium", "heavy"
vbwProfiler.vbwExecuteLine 5751
       If component = WaterCannon Then
vbwProfiler.vbwExecuteLine 5752
           AddPCLproperty "Type of Ammo", .Ammunitiontype, wdList, "AmmunitionType", "water", "acid", "foam"
       End If
vbwProfiler.vbwExecuteLine 5753 'B
vbwProfiler.vbwExecuteLine 5754
       AddPCLproperty "# of Shots", .Shots, wdNumber, "Shots"
vbwProfiler.vbwExecuteLine 5755
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 5756
       AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 5757
       AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5758
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5759
       AddPCLproperty "Type Damage", .TypeDamage, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5760
       AddPCLproperty "Damage", .Damage, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5761
       AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5762
       AddPCLproperty "Max Range (yards)", .MaxRange, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5763
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5764
       AddPCLproperty "Snap Shot", .SnapShot, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5765
       AddPCLproperty "Rate of Fire", .rof, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5766
       AddPCLproperty "# of Shots", .Shots, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5767
       AddPCLproperty "Cost Per Shot", .CPS, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5768
       AddPCLproperty "Weight Per Shot", .WPS, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5769
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5770
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5771
       AddPCLproperty "Loaders", .Loaders, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5772
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5773
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5774
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

'vbwLine 5775:    Case RevolverLauncher, DisposableLauncher, MuzzleloadingLauncher, BreechloadingLauncher, ManualRepeaterLauncher, SlowAutoLoaderLauncher, FastAutoLoaderLauncher, lightAutomaticLauncher, HeavyAutomaticLauncher
    Case IIf(vbwProfiler.vbwExecuteLine(5775), VBWPROFILER_EMPTY, _
        RevolverLauncher), DisposableLauncher, MuzzleloadingLauncher, BreechloadingLauncher, ManualRepeaterLauncher, SlowAutoLoaderLauncher, FastAutoLoaderLauncher, lightAutomaticLauncher, HeavyAutomaticLauncher
vbwProfiler.vbwExecuteLine 5776
       AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5777
       AddPCLproperty "Tech level", .TL, wdList, "TL", 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 5778
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 5779
       AddPCLproperty "Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)"
vbwProfiler.vbwExecuteLine 5780
       AddPCLproperty "Mounting", .Mount, wdList, "Mount", "normal", "concealed"
vbwProfiler.vbwExecuteLine 5781
       AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
vbwProfiler.vbwExecuteLine 5782
       AddPCLproperty "Diameter", .Diameter, wdDouble, "Diameter"
vbwProfiler.vbwExecuteLine 5783
       AddPCLproperty "Maximum Load (lbs)", .MaxLoad, wdDouble, "MaxLoad"
vbwProfiler.vbwExecuteLine 5784
       Select Case component
'vbwLine 5785:        Case RevolverLauncher, DisposableLauncher, MuzzleloadingLauncher, BreechloadingLauncher
        Case IIf(vbwProfiler.vbwExecuteLine(5785), VBWPROFILER_EMPTY, _
        RevolverLauncher), DisposableLauncher, MuzzleloadingLauncher, BreechloadingLauncher
vbwProfiler.vbwExecuteLine 5786
            AddPCLproperty "# of Tubes", .Cylinders, wdNumber, "Cylinders"
       End Select
vbwProfiler.vbwExecuteLine 5787 'B
vbwProfiler.vbwExecuteLine 5788
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 5789
       AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 5790
       AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5791
       AddPCLproperty "Snap Shot", .SnapShot, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5792
       AddPCLproperty "Rate of Fire", .rof, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5793
       AddPCLproperty "# of Shots", .Shots, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5794
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5795
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5796
       AddPCLproperty "Loaders", .Loaders, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5797
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5798
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5799
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

End Select
vbwProfiler.vbwExecuteLine 5800 'B
vbwProfiler.vbwExecuteLine 5801
End With

vbwProfiler.vbwProcOut 272
vbwProfiler.vbwExecuteLine 5802
End Sub

Private Sub ShowPropsForWeaponry2(ByVal component As Integer, ByVal Key As String)
vbwProfiler.vbwProcIn 273
Dim listarray() As String

vbwProfiler.vbwExecuteLine 5803
With m_oCurrentVeh.Components(Key)

' Fill the window with properties for the correct Collection item
vbwProfiler.vbwExecuteLine 5804
Select Case component

'vbwLine 5805:    Case StoneThrower, BoltThrower
    Case IIf(vbwProfiler.vbwExecuteLine(5805), VBWPROFILER_EMPTY, _
        StoneThrower), BoltThrower
vbwProfiler.vbwExecuteLine 5806
       AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5807
       AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 5808
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 5809
       AddPCLproperty "Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)"
vbwProfiler.vbwExecuteLine 5810
       AddPCLproperty "Mounting", .Mount, wdList, "Mount", "normal", "concealed"
vbwProfiler.vbwExecuteLine 5811
       AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
vbwProfiler.vbwExecuteLine 5812
       If component = StoneThrower Then
vbwProfiler.vbwExecuteLine 5813
            AddPCLproperty "Mechanism", .Mechanism, wdList, "Mechanism", "spring-powered", "torsion-powered", "counterweight"
       Else
vbwProfiler.vbwExecuteLine 5814 'B
vbwProfiler.vbwExecuteLine 5815
            AddPCLproperty "Mechanism", .Mechanism, wdList, "Mechanism", "spring-powered", "torsion-powered"
       End If
vbwProfiler.vbwExecuteLine 5816 'B
vbwProfiler.vbwExecuteLine 5817
       AddPCLproperty "Strength", .Strength, wdNumber, "Strength"
vbwProfiler.vbwExecuteLine 5818
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 5819
       AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 5820
       AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5821
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5822
       AddPCLproperty "Type Damage", .TypeDamage, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5823
       AddPCLproperty "Damage", .Damage, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5824
       AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5825
       AddPCLproperty "Max Range (yards)", .MaxRange, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5826
       AddPCLproperty "Min Range (yards)", .MinRange, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5827
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5828
       AddPCLproperty "Snap Shot", .SnapShot, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5829
       AddPCLproperty "Rate of Fire", .rof, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5830
       AddPCLproperty "# of Shots", .Shots, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5831
       AddPCLproperty "Reqt. Loaders", .Loaders, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5832
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5833
       AddPCLproperty "Cost Per Shot", "$" & Format(.CPS, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5834
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5835
       AddPCLproperty "Weight Per Shot", Format(.WPS, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5836
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5837
       AddPCLproperty "Volume Per Shot", Format(.VPS, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5838
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5839
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

'vbwLine 5840:    Case RepeatingBoltThrower
    Case IIf(vbwProfiler.vbwExecuteLine(5840), VBWPROFILER_EMPTY, _
        RepeatingBoltThrower)
vbwProfiler.vbwExecuteLine 5841
       AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5842
       AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 5843
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 5844
       AddPCLproperty "Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)"
vbwProfiler.vbwExecuteLine 5845
       AddPCLproperty "Mounting", .Mount, wdList, "Mount", "normal", "concealed"
vbwProfiler.vbwExecuteLine 5846
       AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
vbwProfiler.vbwExecuteLine 5847
       AddPCLproperty "Mechanism", .Mechanism, wdList, "Mechanism", "spring-powered", "torsion-powered"
vbwProfiler.vbwExecuteLine 5848
       AddPCLproperty "Strength", .Strength, wdNumber, "Strength"
vbwProfiler.vbwExecuteLine 5849
       AddPCLproperty "Magazine Capacity", .MagazineCapacity, wdNumber, "MagazineCapacity"
vbwProfiler.vbwExecuteLine 5850
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 5851
       AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 5852
       AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5853
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5854
       AddPCLproperty "Type Damage", .TypeDamage, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5855
       AddPCLproperty "Damage", .Damage, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5856
       AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5857
       AddPCLproperty "Max Range (yards)", .MaxRange, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5858
       AddPCLproperty "Min Range (yards)", .MinRange, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5859
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5860
       AddPCLproperty "Snap Shot", .SnapShot, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5861
       AddPCLproperty "Rate of Fire", .rof, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5862
       AddPCLproperty "# of Shots", .Shots, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5863
       AddPCLproperty "Reqt. Loaders", .Loaders, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5864
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5865
       AddPCLproperty "Cost Per Shot", "$" & Format(.CPS, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5866
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5867
       AddPCLproperty "Weight Per Shot", Format(.WPS, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5868
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5869
       AddPCLproperty "Volume Per Shot", Format(.VPS, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5870
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5871
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"



'vbwLine 5872:    Case MuzzleLoader, BreechLoader
    Case IIf(vbwProfiler.vbwExecuteLine(5872), VBWPROFILER_EMPTY, _
        MuzzleLoader), BreechLoader
vbwProfiler.vbwExecuteLine 5873
       AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5874
       AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 5875
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 5876
       AddPCLproperty "Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)"
vbwProfiler.vbwExecuteLine 5877
       AddPCLproperty "Mounting", .Mount, wdList, "Mount", "normal", "concealed"
vbwProfiler.vbwExecuteLine 5878
       AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
vbwProfiler.vbwExecuteLine 5879
       AddPCLproperty "Bore Size", .BoreSize, wdDouble, "BoreSize"
vbwProfiler.vbwExecuteLine 5880
       AddPCLproperty "Technology", .technology, wdList, "Technology", "conventional smoothbore", "conventional rifled", "electromag", "gravitic"
vbwProfiler.vbwExecuteLine 5881
       listarray = .FillAmmunitionList
vbwProfiler.vbwExecuteLine 5882
       AddPCLproperty "Ammunition", .Ammunitiontype, wdList, "AmmunitionType", listarray
vbwProfiler.vbwExecuteLine 5883
       AddPCLproperty "Power Option", .PowerOption, wdList, "PowerOption", "normal", "low-powered", "extra-low-powered"
vbwProfiler.vbwExecuteLine 5884
       AddPCLproperty "Recoiless", .Recoiless, wdBool, "Recoiless"
vbwProfiler.vbwExecuteLine 5885
       AddPCLproperty "Barrel Length", .Barrel, wdList, "Barrel", "extremely short", "very short", "short", "medium", "long", "very long", "extremely long"
vbwProfiler.vbwExecuteLine 5886
       AddPCLproperty "# of Fixed Barrels", .Cylinders, wdList, "Cylinders", "1", "2", "3", "4", "5", "6", "7"
       'advanced option not available for unconventinal weapons
vbwProfiler.vbwExecuteLine 5887
       If (.technology = "conventional smoothbore") Or (.technology = "conventional rifled") Then
vbwProfiler.vbwExecuteLine 5888
            AddPCLproperty "Advanced Option", .advancedoption, wdList, "AdvancedOption", "none", "plastic-cased ammunition", "caseless", "liquid propellant", "electrothermal"
       Else
vbwProfiler.vbwExecuteLine 5889 'B
vbwProfiler.vbwExecuteLine 5890
            .advancedoption = "none"
       End If
vbwProfiler.vbwExecuteLine 5891 'B

vbwProfiler.vbwExecuteLine 5892
       AddPCLproperty "Reputation for Quality", .Reliable, wdBool, "Reliable"
vbwProfiler.vbwExecuteLine 5893
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 5894
       AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 5895
       AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5896
        If (component = ElectricGatling) Or (.technology = "electromag") Or (.Electric) Then
vbwProfiler.vbwExecuteLine 5897
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 5898 'B
vbwProfiler.vbwExecuteLine 5899
        If component = MuzzleLoader Then
vbwProfiler.vbwExecuteLine 5900
           AddPCLproperty "Carriage Required", .Carriage, wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 5901 'B
vbwProfiler.vbwExecuteLine 5902
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5903
       If .BurstRadius <> -1 Then
vbwProfiler.vbwExecuteLine 5904
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
vbwProfiler.vbwExecuteLine 5905 'B
vbwProfiler.vbwExecuteLine 5906
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5907
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5908
        If .TypeDamage2 <> "none" Then
vbwProfiler.vbwExecuteLine 5909
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5910
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 5911 'B
vbwProfiler.vbwExecuteLine 5912
       AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5913
       AddPCLproperty "Max Range (yards)", .MaxRange, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5914
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5915
       AddPCLproperty "Snap Shot", .SnapShot, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5916
       AddPCLproperty "Rate of Fire", .sRoF, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5917
       AddPCLproperty "# of Shots", .Shots, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5918
       AddPCLproperty "Reqt. Loaders", .Loaders, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5919
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5920
       AddPCLproperty "Cost Per Shot", "$" & Format(.CPS, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5921
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5922
       AddPCLproperty "Weight Per Shot", Format(.WPS, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5923
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5924
       AddPCLproperty "Volume Per Shot", Format(.VPS, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5925
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5926
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 5927:    Case ManualRepeater
    Case IIf(vbwProfiler.vbwExecuteLine(5927), VBWPROFILER_EMPTY, _
        ManualRepeater)
vbwProfiler.vbwExecuteLine 5928
       AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5929
       AddPCLproperty "Tech level", .TL, wdList, "TL", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16"
vbwProfiler.vbwExecuteLine 5930
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 5931
       AddPCLproperty "Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)"
vbwProfiler.vbwExecuteLine 5932
       AddPCLproperty "Mounting", .Mount, wdList, "Mount", "normal", "concealed"
vbwProfiler.vbwExecuteLine 5933
       AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
vbwProfiler.vbwExecuteLine 5934
       AddPCLproperty "Bore Size", .BoreSize, wdDouble, "BoreSize"
vbwProfiler.vbwExecuteLine 5935
       AddPCLproperty "Technology", .technology, wdList, "Technology", "conventional smoothbore", "conventional rifled", "electromag", "gravitic"
vbwProfiler.vbwExecuteLine 5936
       listarray = .FillAmmunitionList
vbwProfiler.vbwExecuteLine 5937
       AddPCLproperty "Ammunition", .Ammunitiontype, wdList, "AmmunitionType", listarray
vbwProfiler.vbwExecuteLine 5938
       AddPCLproperty "Power Option", .PowerOption, wdList, "PowerOption", "normal", "low-powered", "extra-low-powered"
vbwProfiler.vbwExecuteLine 5939
       AddPCLproperty "Recoiless", .Recoiless, wdBool, "Recoiless"
vbwProfiler.vbwExecuteLine 5940
       AddPCLproperty "Barrel Length", .Barrel, wdList, "Barrel", "extremely short", "very short", "short", "medium", "long", "very long", "extremely long"
       'advanced option not available for unconventinal weapons
vbwProfiler.vbwExecuteLine 5941
       If (.technology = "conventional smoothbore") Or (.technology = "conventional rifled") Then
vbwProfiler.vbwExecuteLine 5942
            AddPCLproperty "Advanced Option", .advancedoption, wdList, "AdvancedOption", "none", "plastic-cased ammunition", "caseless", "liquid propellant", "electrothermal"
       Else
vbwProfiler.vbwExecuteLine 5943 'B
vbwProfiler.vbwExecuteLine 5944
            .advancedoption = "none"
       End If
vbwProfiler.vbwExecuteLine 5945 'B
vbwProfiler.vbwExecuteLine 5946
       AddPCLproperty "Box Magazine", .BoxMagazine, wdBool, "BoxMagazine"
vbwProfiler.vbwExecuteLine 5947
       AddPCLproperty "Reputation for Quality", .Reliable, wdBool, "Reliable"
vbwProfiler.vbwExecuteLine 5948
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 5949
       AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 5950
       AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5951
        If (component = ElectricGatling) Or (.technology = "electromag") Or (.Electric) Then
vbwProfiler.vbwExecuteLine 5952
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 5953 'B
vbwProfiler.vbwExecuteLine 5954
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5955
       If .BurstRadius <> -1 Then
vbwProfiler.vbwExecuteLine 5956
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
vbwProfiler.vbwExecuteLine 5957 'B
vbwProfiler.vbwExecuteLine 5958
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5959
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5960
        If .TypeDamage2 <> "none" Then
vbwProfiler.vbwExecuteLine 5961
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5962
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 5963 'B
vbwProfiler.vbwExecuteLine 5964
       AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5965
       AddPCLproperty "Max Range (yards)", .MaxRange, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5966
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5967
       AddPCLproperty "Snap Shot", .SnapShot, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5968
       AddPCLproperty "Rate of Fire", .sRoF, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5969
       AddPCLproperty "# of Shots", .Shots, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5970
       AddPCLproperty "Reqt. Loaders", .Loaders, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5971
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5972
       AddPCLproperty "Cost Per Shot", "$" & Format(.CPS, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5973
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5974
       AddPCLproperty "Weight Per Shot", Format(.WPS, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5975
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5976
       AddPCLproperty "Volume Per Shot", Format(.VPS, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5977
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5978
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 5979:    Case Revolver, MechanicalGatling
    Case IIf(vbwProfiler.vbwExecuteLine(5979), VBWPROFILER_EMPTY, _
        Revolver), MechanicalGatling
            'NOTE: These allow for user modifieable Rates of Fire
    'power only needs to be displayed for elec.gat.
vbwProfiler.vbwExecuteLine 5980
       AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 5981
       AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 5982
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 5983
       AddPCLproperty "Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)"
vbwProfiler.vbwExecuteLine 5984
       AddPCLproperty "Mounting", .Mount, wdList, "Mount", "normal", "concealed"
vbwProfiler.vbwExecuteLine 5985
       AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
vbwProfiler.vbwExecuteLine 5986
       AddPCLproperty "Bore Size", .BoreSize, wdDouble, "BoreSize"
vbwProfiler.vbwExecuteLine 5987
       AddPCLproperty "Technology", .technology, wdList, "Technology", "conventional smoothbore", "conventional rifled", "electromag", "gravitic"
vbwProfiler.vbwExecuteLine 5988
       listarray = .FillAmmunitionList
vbwProfiler.vbwExecuteLine 5989
       AddPCLproperty "Ammunition", .Ammunitiontype, wdList, "AmmunitionType", listarray
vbwProfiler.vbwExecuteLine 5990
       If component = MechanicalGatling Then
vbwProfiler.vbwExecuteLine 5991
           AddPCLproperty "Operator DX + Skill", .DXPlusSkill, wdDouble, "DXPlusSkill"
       End If
vbwProfiler.vbwExecuteLine 5992 'B
vbwProfiler.vbwExecuteLine 5993
       AddPCLproperty "Power Option", .PowerOption, wdList, "PowerOption", "normal", "low-powered", "extra-low-powered"
vbwProfiler.vbwExecuteLine 5994
       AddPCLproperty "Recoiless", .Recoiless, wdBool, "Recoiless"
vbwProfiler.vbwExecuteLine 5995
       AddPCLproperty "Barrel Length", .Barrel, wdList, "Barrel", "extremely short", "very short", "short", "medium", "long", "very long", "extremely long"
vbwProfiler.vbwExecuteLine 5996
        If component = Revolver Then
vbwProfiler.vbwExecuteLine 5997
           AddPCLproperty "# of Cylinders", .Cylinders, wdList, "Cylinders", "3", "4", "5", "6", "7"
        Else
vbwProfiler.vbwExecuteLine 5998 'B
vbwProfiler.vbwExecuteLine 5999
           AddPCLproperty "# of Barrels", .Cylinders, wdList, "Cylinders", "3", "4", "5", "6", "7"
        End If
vbwProfiler.vbwExecuteLine 6000 'B
       'advanced option not available for unconventinal weapons
vbwProfiler.vbwExecuteLine 6001
       If (.technology = "conventional smoothbore") Or (.technology = "conventional rifled") Then
vbwProfiler.vbwExecuteLine 6002
            AddPCLproperty "Advanced Option", .advancedoption, wdList, "AdvancedOption", "none", "plastic-cased ammunition", "caseless", "liquid propellant", "electrothermal"
       Else
vbwProfiler.vbwExecuteLine 6003 'B
vbwProfiler.vbwExecuteLine 6004
            .advancedoption = "none"
       End If
vbwProfiler.vbwExecuteLine 6005 'B
vbwProfiler.vbwExecuteLine 6006
       AddPCLproperty "Reputation for Quality", .Reliable, wdBool, "Reliable"
vbwProfiler.vbwExecuteLine 6007
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 6008
       AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 6009
       AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6010
        If (component = ElectricGatling) Or (.technology = "electromag") Or (.Electric) Then
vbwProfiler.vbwExecuteLine 6011
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 6012 'B
vbwProfiler.vbwExecuteLine 6013
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6014
       If .BurstRadius <> -1 Then
vbwProfiler.vbwExecuteLine 6015
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
vbwProfiler.vbwExecuteLine 6016 'B
vbwProfiler.vbwExecuteLine 6017
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6018
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6019
        If .TypeDamage2 <> "none" Then
vbwProfiler.vbwExecuteLine 6020
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6021
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 6022 'B
vbwProfiler.vbwExecuteLine 6023
       AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6024
       AddPCLproperty "Max Range (yards)", .MaxRange, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6025
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6026
       AddPCLproperty "Snap Shot", .SnapShot, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6027
       AddPCLproperty "Rate of Fire", .sRoF, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6028
       AddPCLproperty "# of Shots", .Shots, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6029
       AddPCLproperty "Reqt. Loaders", .Loaders, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6030
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6031
       AddPCLproperty "Cost Per Shot", "$" & Format(.CPS, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6032
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6033
       AddPCLproperty "Weight Per Shot", Format(.WPS, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6034
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6035
       AddPCLproperty "Volume Per Shot", Format(.VPS, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6036
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6037
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

'vbwLine 6038:    Case ElectricGatling
    Case IIf(vbwProfiler.vbwExecuteLine(6038), VBWPROFILER_EMPTY, _
        ElectricGatling)
    'NOTE: This allows for user modifieable Rate of Fire
vbwProfiler.vbwExecuteLine 6039
       AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6040
       AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6041
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6042
       AddPCLproperty "Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)"
vbwProfiler.vbwExecuteLine 6043
       AddPCLproperty "Mounting", .Mount, wdList, "Mount", "normal", "concealed"
vbwProfiler.vbwExecuteLine 6044
       AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
vbwProfiler.vbwExecuteLine 6045
       AddPCLproperty "Bore Size", .BoreSize, wdDouble, "BoreSize"
vbwProfiler.vbwExecuteLine 6046
       AddPCLproperty "Technology", .technology, wdList, "Technology", "conventional smoothbore", "conventional rifled", "electromag", "gravitic"
vbwProfiler.vbwExecuteLine 6047
       listarray = .FillAmmunitionList
vbwProfiler.vbwExecuteLine 6048
       AddPCLproperty "Ammunition", .Ammunitiontype, wdList, "AmmunitionType", listarray
vbwProfiler.vbwExecuteLine 6049
       listarray = .FillRoFList
vbwProfiler.vbwExecuteLine 6050
       AddPCLproperty "Rate of Fire", .dRoF, wdList, "dRoF", listarray
vbwProfiler.vbwExecuteLine 6051
       AddPCLproperty "Power Option", .PowerOption, wdList, "PowerOption", "normal", "low-powered", "extra-low-powered"
vbwProfiler.vbwExecuteLine 6052
       AddPCLproperty "Recoiless", .Recoiless, wdBool, "Recoiless"
vbwProfiler.vbwExecuteLine 6053
       AddPCLproperty "Barrel Length", .Barrel, wdList, "Barrel", "extremely short", "very short", "short", "medium", "long", "very long", "extremely long"
vbwProfiler.vbwExecuteLine 6054
        If component = Revolver Then
vbwProfiler.vbwExecuteLine 6055
           AddPCLproperty "# of Cylinders", .Cylinders, wdList, "Cylinders", "3", "4", "5", "6", "7"
        Else
vbwProfiler.vbwExecuteLine 6056 'B
vbwProfiler.vbwExecuteLine 6057
           AddPCLproperty "# of Barrels", .Cylinders, wdList, "Cylinders", "3", "4", "5", "6", "7"
        End If
vbwProfiler.vbwExecuteLine 6058 'B
       'advanced option not available for unconventinal weapons
vbwProfiler.vbwExecuteLine 6059
       If (.technology = "conventional smoothbore") Or (.technology = "conventional rifled") Then
vbwProfiler.vbwExecuteLine 6060
            AddPCLproperty "Advanced Option", .advancedoption, wdList, "AdvancedOption", "none", "plastic-cased ammunition", "caseless", "liquid propellant", "electrothermal"
       Else
vbwProfiler.vbwExecuteLine 6061 'B
vbwProfiler.vbwExecuteLine 6062
            .advancedoption = "none"
       End If
vbwProfiler.vbwExecuteLine 6063 'B
vbwProfiler.vbwExecuteLine 6064
       AddPCLproperty "Reputation for Quality", .Reliable, wdBool, "Reliable"
vbwProfiler.vbwExecuteLine 6065
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 6066
       AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 6067
       AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6068
        If (component = ElectricGatling) Or (.technology = "electromag") Or (.Electric) Then
vbwProfiler.vbwExecuteLine 6069
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 6070 'B
vbwProfiler.vbwExecuteLine 6071
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6072
       If .BurstRadius <> -1 Then
vbwProfiler.vbwExecuteLine 6073
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
vbwProfiler.vbwExecuteLine 6074 'B
vbwProfiler.vbwExecuteLine 6075
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6076
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6077
        If .TypeDamage2 <> "none" Then
vbwProfiler.vbwExecuteLine 6078
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6079
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 6080 'B
vbwProfiler.vbwExecuteLine 6081
       AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6082
       AddPCLproperty "Max Range (yards)", .MaxRange, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6083
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6084
       AddPCLproperty "Snap Shot", .SnapShot, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6085
       AddPCLproperty "# of Shots", .Shots, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6086
       AddPCLproperty "Reqt. Loaders", .Loaders, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6087
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6088
       AddPCLproperty "Cost Per Shot", "$" & Format(.CPS, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6089
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6090
       AddPCLproperty "Weight Per Shot", Format(.WPS, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6091
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6092
       AddPCLproperty "Volume Per Shot", Format(.VPS, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6093
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6094
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6095:    Case SlowAutoloader, FastAutoloader
    Case IIf(vbwProfiler.vbwExecuteLine(6095), VBWPROFILER_EMPTY, _
        SlowAutoloader), FastAutoloader
vbwProfiler.vbwExecuteLine 6096
        AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6097
       AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6098
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6099
       AddPCLproperty "Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)"
vbwProfiler.vbwExecuteLine 6100
       AddPCLproperty "Mounting", .Mount, wdList, "Mount", "normal", "concealed"
vbwProfiler.vbwExecuteLine 6101
       AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
vbwProfiler.vbwExecuteLine 6102
       AddPCLproperty "Bore Size", .BoreSize, wdDouble, "BoreSize"
vbwProfiler.vbwExecuteLine 6103
       AddPCLproperty "Technology", .technology, wdList, "Technology", "conventional smoothbore", "conventional rifled", "electromag", "gravitic"
vbwProfiler.vbwExecuteLine 6104
       listarray = .FillAmmunitionList
vbwProfiler.vbwExecuteLine 6105
       AddPCLproperty "Ammunition", .Ammunitiontype, wdList, "AmmunitionType", listarray
vbwProfiler.vbwExecuteLine 6106
       AddPCLproperty "Power Option", .PowerOption, wdList, "PowerOption", "normal", "low-powered", "extra-low-powered"
vbwProfiler.vbwExecuteLine 6107
       AddPCLproperty "Recoiless", .Recoiless, wdBool, "Recoiless"
vbwProfiler.vbwExecuteLine 6108
       AddPCLproperty "Barrel Length", .Barrel, wdList, "Barrel", "extremely short", "very short", "short", "medium", "long", "very long", "extremely long"
vbwProfiler.vbwExecuteLine 6109
       AddPCLproperty "Electric Loading", .Electric, wdBool, "Electric"
       'advanced option not available for unconventinal weapons
vbwProfiler.vbwExecuteLine 6110
       If (.technology = "conventional smoothbore") Or (.technology = "conventional rifled") Then
vbwProfiler.vbwExecuteLine 6111
            AddPCLproperty "Advanced Option", .advancedoption, wdList, "AdvancedOption", "none", "plastic-cased ammunition", "caseless", "liquid propellant", "electrothermal"
       Else
vbwProfiler.vbwExecuteLine 6112 'B
vbwProfiler.vbwExecuteLine 6113
            .advancedoption = "none"
       End If
vbwProfiler.vbwExecuteLine 6114 'B
vbwProfiler.vbwExecuteLine 6115
       AddPCLproperty "Reputation for Quality", .Reliable, wdBool, "Reliable"
vbwProfiler.vbwExecuteLine 6116
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 6117
       AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 6118
       AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6119
        If (component = ElectricGatling) Or (.technology = "electromag") Or (.technology = "gravitic") Or (.Electric) Then
vbwProfiler.vbwExecuteLine 6120
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 6121 'B
vbwProfiler.vbwExecuteLine 6122
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6123
       If .BurstRadius <> -1 Then
vbwProfiler.vbwExecuteLine 6124
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
vbwProfiler.vbwExecuteLine 6125 'B
vbwProfiler.vbwExecuteLine 6126
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6127
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6128
        If .TypeDamage2 <> "none" Then
vbwProfiler.vbwExecuteLine 6129
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6130
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 6131 'B
vbwProfiler.vbwExecuteLine 6132
       AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6133
       AddPCLproperty "Max Range (yards)", .MaxRange, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6134
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6135
       AddPCLproperty "Snap Shot", .SnapShot, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6136
       AddPCLproperty "Rate of Fire", .sRoF, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6137
       AddPCLproperty "# of Shots", .Shots, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6138
       AddPCLproperty "Reqt. Loaders", .Loaders, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6139
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6140
       AddPCLproperty "Cost Per Shot", "$" & Format(.CPS, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6141
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6142
       AddPCLproperty "Weight Per Shot", Format(.WPS, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6143
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6144
       AddPCLproperty "Volume Per Shot", Format(.VPS, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6145
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6146
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

'vbwLine 6147:    Case lightAutomatic, HeavyAutomatic
    Case IIf(vbwProfiler.vbwExecuteLine(6147), VBWPROFILER_EMPTY, _
        lightAutomatic), HeavyAutomatic
        'note: these allow for user edit-able Rates of Fire
vbwProfiler.vbwExecuteLine 6148
       AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6149
       AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6150
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6151
       AddPCLproperty "Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)"
vbwProfiler.vbwExecuteLine 6152
       AddPCLproperty "Mounting", .Mount, wdList, "Mount", "normal", "concealed"
vbwProfiler.vbwExecuteLine 6153
       AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
vbwProfiler.vbwExecuteLine 6154
       AddPCLproperty "Bore Size", .BoreSize, wdDouble, "BoreSize"
vbwProfiler.vbwExecuteLine 6155
       AddPCLproperty "Technology", .technology, wdList, "Technology", "conventional smoothbore", "conventional rifled", "electromag", "gravitic"
vbwProfiler.vbwExecuteLine 6156
       listarray = .FillAmmunitionList
vbwProfiler.vbwExecuteLine 6157
       AddPCLproperty "Ammunition", .Ammunitiontype, wdList, "AmmunitionType", listarray
vbwProfiler.vbwExecuteLine 6158
       listarray = .FillRoFList
vbwProfiler.vbwExecuteLine 6159
       AddPCLproperty "Rate of Fire", .dRoF, wdList, "dRoF", listarray
vbwProfiler.vbwExecuteLine 6160
       AddPCLproperty "Power Option", .PowerOption, wdList, "PowerOption", "normal", "low-powered", "extra-low-powered"
vbwProfiler.vbwExecuteLine 6161
       AddPCLproperty "Recoiless", .Recoiless, wdBool, "Recoiless"
vbwProfiler.vbwExecuteLine 6162
       AddPCLproperty "Barrel Length", .Barrel, wdList, "Barrel", "extremely short", "very short", "short", "medium", "long", "very long", "extremely long"
vbwProfiler.vbwExecuteLine 6163
       AddPCLproperty "Electric Loading", .Electric, wdBool, "Electric"
       'advanced option not available for unconventinal weapons
vbwProfiler.vbwExecuteLine 6164
       If (.technology = "conventional smoothbore") Or (.technology = "conventional rifled") Then
vbwProfiler.vbwExecuteLine 6165
            AddPCLproperty "Advanced Option", .advancedoption, wdList, "AdvancedOption", "none", "plastic-cased ammunition", "caseless", "liquid propellant", "electrothermal"
       Else
vbwProfiler.vbwExecuteLine 6166 'B
vbwProfiler.vbwExecuteLine 6167
            .advancedoption = "none"
       End If
vbwProfiler.vbwExecuteLine 6168 'B
vbwProfiler.vbwExecuteLine 6169
       AddPCLproperty "Reputation for Quality", .Reliable, wdBool, "Reliable"
vbwProfiler.vbwExecuteLine 6170
       AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 6171
       AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 6172
       AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6173
        If (component = ElectricGatling) Or (.technology = "electromag") Or (.Electric) Then
vbwProfiler.vbwExecuteLine 6174
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 6175 'B
vbwProfiler.vbwExecuteLine 6176
       AddPCLproperty "Malfunction", .Malfunction, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6177
       If .BurstRadius <> -1 Then
vbwProfiler.vbwExecuteLine 6178
         AddPCLproperty "Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius"
       End If
vbwProfiler.vbwExecuteLine 6179 'B
vbwProfiler.vbwExecuteLine 6180
       AddPCLproperty "Type Damage 1", .TypeDamage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6181
       AddPCLproperty "Damage 1", .Damage1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6182
        If .TypeDamage2 <> "none" Then
vbwProfiler.vbwExecuteLine 6183
           AddPCLproperty "Type Damage 2", .TypeDamage2, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6184
       AddPCLproperty "Damage 2", .Damage2, wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 6185 'B
vbwProfiler.vbwExecuteLine 6186
       AddPCLproperty "1/2 Damage (yards)", .halfDamage, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6187
       AddPCLproperty "Max Range (yards)", .MaxRange, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6188
       AddPCLproperty "Accuracy", .Accuracy, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6189
       AddPCLproperty "Snap Shot", .SnapShot, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6190
       AddPCLproperty "# of Shots", .Shots, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6191
       AddPCLproperty "Reqt. Loaders", .Loaders, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6192
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6193
       AddPCLproperty "Cost Per Shot", "$" & Format(.CPS, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6194
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6195
       AddPCLproperty "Weight Per Shot", Format(.WPS, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6196
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6197
       AddPCLproperty "Volume Per Shot", Format(.VPS, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6198
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6199
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

'vbwLine 6200:    Case AntiBlastMagazine
    Case IIf(vbwProfiler.vbwExecuteLine(6200), VBWPROFILER_EMPTY, _
        AntiBlastMagazine)
vbwProfiler.vbwExecuteLine 6201
       AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6202
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6203
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"

'vbwLine 6204:    Case UniversalMount, CasemateMount, DoorMount, Cyberslave, FullStabilizationGear, PartialStabilizationGear
    Case IIf(vbwProfiler.vbwExecuteLine(6204), VBWPROFILER_EMPTY, _
        UniversalMount), CasemateMount, DoorMount, Cyberslave, FullStabilizationGear, PartialStabilizationGear
vbwProfiler.vbwExecuteLine 6205
       AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6206
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6207
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6208
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6209
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6210
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"



'vbwLine 6211:    Case WeaponBay, HardPoint
    Case IIf(vbwProfiler.vbwExecuteLine(6211), VBWPROFILER_EMPTY, _
        WeaponBay), HardPoint
vbwProfiler.vbwExecuteLine 6212
       AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6213
       AddPCLproperty "Index", .index, wdNumber, "Index"
vbwProfiler.vbwExecuteLine 6214
       AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6215
       AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
vbwProfiler.vbwExecuteLine 6216
       AddPCLproperty "Maximum Load (lbs)", .loadcapacity, wdDouble, "LoadCapacity"
vbwProfiler.vbwExecuteLine 6217
       AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6218
       AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
vbwProfiler.vbwExecuteLine 6219
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6220
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6221
        If component = WeaponBay Then
vbwProfiler.vbwExecuteLine 6222
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6223
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6224
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
        End If
vbwProfiler.vbwExecuteLine 6225 'B


'vbwLine 6226:    Case Ammunition
    Case IIf(vbwProfiler.vbwExecuteLine(6226), VBWPROFILER_EMPTY, _
        Ammunition)
vbwProfiler.vbwExecuteLine 6227
       AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6228
       AddPCLproperty "# of Shots", .NumShots, wdNumber, "NumShots"
vbwProfiler.vbwExecuteLine 6229
       AddPCLproperty "Lock Ammo Settings", .Locked, wdBool, "Locked"
vbwProfiler.vbwExecuteLine 6230
       AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6231
       AddPCLproperty "Ammo Type", .Ammunitiontype, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6232
       AddPCLproperty "CPS", "$" & Format(.CPS, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6233
       AddPCLproperty "WPS", Format(.WPS, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6234
       AddPCLproperty "VPS", Format(.VPS, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6235
       AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6236
       AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6237
       AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6238
       AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6239
       AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


End Select
vbwProfiler.vbwExecuteLine 6240 'B

vbwProfiler.vbwExecuteLine 6241
End With
vbwProfiler.vbwProcOut 273
vbwProfiler.vbwExecuteLine 6242
End Sub

Private Sub ShowPropsForBody()
vbwProfiler.vbwProcIn 274
vbwProfiler.vbwExecuteLine 6243
    With m_oCurrentVeh.Components(BODY_KEY)
vbwProfiler.vbwExecuteLine 6244
        AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6245
        AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6246
        AddPCLproperty "Compartmentalization", .Compartmentalization, wdList, "Compartmentalization", "none", "heavy", "total"
vbwProfiler.vbwExecuteLine 6247
        AddPCLproperty "Flexibody Option", .FlexibodyOption, wdBool, "FlexibodyOption"
vbwProfiler.vbwExecuteLine 6248
        AddPCLproperty "Improved Flexibody Suspension", .ImprovedSuspension, wdBool, "ImprovedSuspension"
vbwProfiler.vbwExecuteLine 6249
        AddPCLproperty "Lifting Body", .liftingbody, wdBool, "LiftingBody"
vbwProfiler.vbwExecuteLine 6250
        AddPCLproperty "Top Deck", .TopDeck, wdBool, "Topdeck"
vbwProfiler.vbwExecuteLine 6251
        AddPCLproperty "% Covered Deck", .PercentCovered, wdNumber, "PercentCovered"
vbwProfiler.vbwExecuteLine 6252
        AddPCLproperty "% Flight Deck", .PercentFlightDeck, wdNumber, "PercentFlightDeck"
vbwProfiler.vbwExecuteLine 6253
        AddPCLproperty "Flight Deck Option", .flightdeckoption, wdList, "flightdeckoption", "none", "landing pad", "angled flight deck"
vbwProfiler.vbwExecuteLine 6254
        AddPCLproperty "Slope Right", .SlopeR, wdList, "sloper", "none", "30 degrees", "60 degrees"
vbwProfiler.vbwExecuteLine 6255
        AddPCLproperty "Slope Left", .slopel, wdList, "slopel", "none", "30 degrees", "60 degrees"
vbwProfiler.vbwExecuteLine 6256
        AddPCLproperty "Slope Front", .slopef, wdList, "slopeF", "none", "30 degrees", "60 degrees"
vbwProfiler.vbwExecuteLine 6257
        AddPCLproperty "Slope Back", .slopeb, wdList, "slopeb", "none", "30 degrees", "60 degrees"
vbwProfiler.vbwExecuteLine 6258
        AddPCLproperty "Empty Space", .EmptySpace, wdDouble, "EmptySpace"
vbwProfiler.vbwExecuteLine 6259
        AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
vbwProfiler.vbwExecuteLine 6260
        AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
vbwProfiler.vbwExecuteLine 6261
        AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
vbwProfiler.vbwExecuteLine 6262
        AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
vbwProfiler.vbwExecuteLine 6263
        AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
vbwProfiler.vbwExecuteLine 6264
        AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"

vbwProfiler.vbwExecuteLine 6265
        AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6266
        AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
vbwProfiler.vbwExecuteLine 6267
        AddPCLproperty "Top Deck Area", Format(.TotalDeckArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6268
        AddPCLproperty "Flight Deck Length", Format(.flightdecklength, "standard") & " ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6269
        AddPCLproperty "Flight Deck Area", Format(.FlightDeckArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6270
        AddPCLproperty "Covered Deck Area", Format(.covereddeckarea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6271
        AddPCLproperty "Deck Cost", "$" & Format(.DeckCost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6272
        AddPCLproperty "Deck Weight", Format(.DeckWeight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6273
        AddPCLproperty "Body Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6274
        AddPCLproperty "Body Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6275
        AddPCLproperty "Body Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6276
        AddPCLproperty "Body Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6277
        AddPCLproperty "Access Space", Format(.AccessSpace, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6278
        AddPCLproperty "Minimum Volume", .MinimumVolume, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6279
        AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6280
    End With

vbwProfiler.vbwProcOut 274
vbwProfiler.vbwExecuteLine 6281
End Sub
Private Sub ShowPropsForSubAssemblies(ByVal component As Integer, ByVal Key As String)
vbwProfiler.vbwProcIn 275

vbwProfiler.vbwExecuteLine 6282
With m_oCurrentVeh.Components(Key)

' Fill the window with properties for the correct Collection item
vbwProfiler.vbwExecuteLine 6283
Select Case component


'vbwLine 6284:        Case Wheel
        Case IIf(vbwProfiler.vbwExecuteLine(6284), VBWPROFILER_EMPTY, _
        Wheel)
vbwProfiler.vbwExecuteLine 6285
            AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6286
            AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6287
            AddPCLproperty "Wheel Type", .subtype, wdList, "Subtype", "standard", "small", "heavy", "railway", "off-road", "retractable"
vbwProfiler.vbwExecuteLine 6288
            AddPCLproperty "# of Wheels", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6289
            AddPCLproperty "Improved Suspension", .ImprovedSuspension, wdBool, "ImprovedSuspension"
vbwProfiler.vbwExecuteLine 6290
            AddPCLproperty "Retract Location", .RetractLocation, wdList, "RetractLocation", "none", "body", "body & wings"
vbwProfiler.vbwExecuteLine 6291
            AddPCLproperty "Wheel Blades", .Wheelblades, wdList, "Wheelblades", "none", "fixed", "rectractable"
vbwProfiler.vbwExecuteLine 6292
            AddPCLproperty "Snow Tires", .snowtires, wdBool, "Snowtires"
vbwProfiler.vbwExecuteLine 6293
            AddPCLproperty "Racing Tires", .racingtires, wdBool, "RacingTires"
vbwProfiler.vbwExecuteLine 6294
            AddPCLproperty "Puncture Resistant", .PunctureResistant, wdBool, "PunctureResistant"
vbwProfiler.vbwExecuteLine 6295
            AddPCLproperty "Improved Brakes", .ImprovedBrakes, wdBool, "ImprovedBrakes"
vbwProfiler.vbwExecuteLine 6296
            AddPCLproperty "All Wheel Steering", .AllwheelSteering, wdBool, "AllWheelSteering"
vbwProfiler.vbwExecuteLine 6297
            AddPCLproperty "Smart Wheels", .Smartwheels, wdBool, "SmartWheels"
vbwProfiler.vbwExecuteLine 6298
            AddPCLproperty "Empty Space", .EmptySpace, wdDouble, "EmptySpace"
            'note, no empty space allowed
vbwProfiler.vbwExecuteLine 6299
            AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
vbwProfiler.vbwExecuteLine 6300
            AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
vbwProfiler.vbwExecuteLine 6301
            AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
vbwProfiler.vbwExecuteLine 6302
            AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
vbwProfiler.vbwExecuteLine 6303
            AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
vbwProfiler.vbwExecuteLine 6304
            AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
vbwProfiler.vbwExecuteLine 6305
            AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6306
            AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
vbwProfiler.vbwExecuteLine 6307
            AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6308
            AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6309
            AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6310
            AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6311
            AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

'vbwLine 6312:    Case Skid
    Case IIf(vbwProfiler.vbwExecuteLine(6312), VBWPROFILER_EMPTY, _
        Skid)
vbwProfiler.vbwExecuteLine 6313
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6314
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6315
           AddPCLproperty "# of Skids", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6316
           AddPCLproperty "Improved Suspension", .ImprovedSuspension, wdBool, "ImprovedSuspension"
vbwProfiler.vbwExecuteLine 6317
           AddPCLproperty "Retract Location", .RetractLocation, wdList, "RetractLocation", "none", "body", "body & wings"
            'note, no empty space allowed
vbwProfiler.vbwExecuteLine 6318
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
vbwProfiler.vbwExecuteLine 6319
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
vbwProfiler.vbwExecuteLine 6320
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
vbwProfiler.vbwExecuteLine 6321
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
vbwProfiler.vbwExecuteLine 6322
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
vbwProfiler.vbwExecuteLine 6323
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
vbwProfiler.vbwExecuteLine 6324
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6325
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
vbwProfiler.vbwExecuteLine 6326
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6327
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6328
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6329
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6330
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6331:        Case Track
        Case IIf(vbwProfiler.vbwExecuteLine(6331), VBWPROFILER_EMPTY, _
        Track)
vbwProfiler.vbwExecuteLine 6332
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6333
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6334
           AddPCLproperty "Track Type", .subtype, wdList, "SubType", "tracks", "halftracks", "skitracks"
vbwProfiler.vbwExecuteLine 6335
           AddPCLproperty "# of Tracks", .Quantity, wdList, "Quantity", 2, 4
vbwProfiler.vbwExecuteLine 6336
           AddPCLproperty "Improved Suspension", .ImprovedSuspension, wdBool, "ImprovedSuspension"
            'note, no empty space allowed
vbwProfiler.vbwExecuteLine 6337
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
vbwProfiler.vbwExecuteLine 6338
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
vbwProfiler.vbwExecuteLine 6339
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
vbwProfiler.vbwExecuteLine 6340
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
vbwProfiler.vbwExecuteLine 6341
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
vbwProfiler.vbwExecuteLine 6342
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
vbwProfiler.vbwExecuteLine 6343
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6344
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
vbwProfiler.vbwExecuteLine 6345
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6346
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6347
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6348
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6349
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6350:        Case Arm
        Case IIf(vbwProfiler.vbwExecuteLine(6350), VBWPROFILER_EMPTY, _
        Arm)
vbwProfiler.vbwExecuteLine 6351
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6352
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6353
           AddPCLproperty "Index", .index, wdNumber, "Index"
vbwProfiler.vbwExecuteLine 6354
           AddPCLproperty "Orientation", .Orientation, wdList, "Orientation", "top", "underside", "front", "back", "left", "right"
vbwProfiler.vbwExecuteLine 6355
           AddPCLproperty "Empty Space", .EmptySpace, wdDouble, "EmptySpace"
vbwProfiler.vbwExecuteLine 6356
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
vbwProfiler.vbwExecuteLine 6357
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
vbwProfiler.vbwExecuteLine 6358
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
vbwProfiler.vbwExecuteLine 6359
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
vbwProfiler.vbwExecuteLine 6360
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
vbwProfiler.vbwExecuteLine 6361
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
vbwProfiler.vbwExecuteLine 6362
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6363
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
vbwProfiler.vbwExecuteLine 6364
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6365
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6366
           AddPCLproperty "Access Space", Format(.AccessSpace, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6367
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6368
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6369
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6370:        Case Leg
        Case IIf(vbwProfiler.vbwExecuteLine(6370), VBWPROFILER_EMPTY, _
        Leg)
vbwProfiler.vbwExecuteLine 6371
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6372
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6373
           AddPCLproperty "Index", .index, wdNumber, "Index"
vbwProfiler.vbwExecuteLine 6374
           AddPCLproperty "Improved Suspension", .ImprovedSuspension, wdBool, "ImprovedSuspension"
vbwProfiler.vbwExecuteLine 6375
           AddPCLproperty "Empty Space", .EmptySpace, wdDouble, "EmptySpace"
vbwProfiler.vbwExecuteLine 6376
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
vbwProfiler.vbwExecuteLine 6377
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
vbwProfiler.vbwExecuteLine 6378
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
vbwProfiler.vbwExecuteLine 6379
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
vbwProfiler.vbwExecuteLine 6380
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
vbwProfiler.vbwExecuteLine 6381
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
vbwProfiler.vbwExecuteLine 6382
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6383
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
vbwProfiler.vbwExecuteLine 6384
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6385
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6386
           AddPCLproperty "Access Space", Format(.AccessSpace, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6387
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6388
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6389
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6390:        Case Wing
        Case IIf(vbwProfiler.vbwExecuteLine(6390), VBWPROFILER_EMPTY, _
        Wing)
vbwProfiler.vbwExecuteLine 6391
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6392
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6393
           AddPCLproperty "Index", .index, wdNumber, "Index"
vbwProfiler.vbwExecuteLine 6394
           AddPCLproperty "Orientation", .Orientation, wdList, "Orientation", "left", "right"
vbwProfiler.vbwExecuteLine 6395
           AddPCLproperty "Wing Type", .subtype, wdList, "SubType", "standard", "STOL", "biplane", "triplane", "high agility", "flarecraft", "stub"
vbwProfiler.vbwExecuteLine 6396
           AddPCLproperty "Controlled Instability", .ControlledInstability, wdBool, "ControlledInstability"
vbwProfiler.vbwExecuteLine 6397
           AddPCLproperty "Folding Wings", .Folding, wdBool, "Folding"
vbwProfiler.vbwExecuteLine 6398
           AddPCLproperty "Variable Sweep Wings", .VariableSweep, wdList, "VariableSweep", "none", "manual", "automatic"
vbwProfiler.vbwExecuteLine 6399
           AddPCLproperty "Empty Space", .EmptySpace, wdDouble, "EmptySpace"
vbwProfiler.vbwExecuteLine 6400
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
vbwProfiler.vbwExecuteLine 6401
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
vbwProfiler.vbwExecuteLine 6402
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
vbwProfiler.vbwExecuteLine 6403
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
vbwProfiler.vbwExecuteLine 6404
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
vbwProfiler.vbwExecuteLine 6405
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
vbwProfiler.vbwExecuteLine 6406
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6407
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
vbwProfiler.vbwExecuteLine 6408
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6409
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6410
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6411
           AddPCLproperty "Access Space", Format(.AccessSpace, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6412
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6413
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6414:        Case AutogyroRotor, TTRotor, CARotor, MMRotor
        Case IIf(vbwProfiler.vbwExecuteLine(6414), VBWPROFILER_EMPTY, _
        AutogyroRotor), TTRotor, CARotor, MMRotor
vbwProfiler.vbwExecuteLine 6415
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6416
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6417
           AddPCLproperty "Index", .index, wdNumber, "Index"
vbwProfiler.vbwExecuteLine 6418
           AddPCLproperty "Controlled Instability", .ControlledInstability, wdBool, "ControlledInstability"
vbwProfiler.vbwExecuteLine 6419
           AddPCLproperty "Folding Rotors", .Folding, wdBool, "Folding"
            'note, no empty space allowed
vbwProfiler.vbwExecuteLine 6420
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
vbwProfiler.vbwExecuteLine 6421
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
vbwProfiler.vbwExecuteLine 6422
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
vbwProfiler.vbwExecuteLine 6423
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
vbwProfiler.vbwExecuteLine 6424
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
vbwProfiler.vbwExecuteLine 6425
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
vbwProfiler.vbwExecuteLine 6426
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6427
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
vbwProfiler.vbwExecuteLine 6428
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6429
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6430
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6431
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6432
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6433:        Case Hydrofoil
        Case IIf(vbwProfiler.vbwExecuteLine(6433), VBWPROFILER_EMPTY, _
        Hydrofoil)
vbwProfiler.vbwExecuteLine 6434
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6435
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6436
           AddPCLproperty "Index", .index, wdNumber, "Index"
            'note, no empty space allowed
vbwProfiler.vbwExecuteLine 6437
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
vbwProfiler.vbwExecuteLine 6438
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
vbwProfiler.vbwExecuteLine 6439
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
vbwProfiler.vbwExecuteLine 6440
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
vbwProfiler.vbwExecuteLine 6441
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
vbwProfiler.vbwExecuteLine 6442
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
vbwProfiler.vbwExecuteLine 6443
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6444
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
vbwProfiler.vbwExecuteLine 6445
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6446
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
            'access space because Aquatic propulsion can be placed in them
vbwProfiler.vbwExecuteLine 6447
           AddPCLproperty "Access Space", Format(.AccessSpace, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6448
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6449
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6450
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6451:        Case Hovercraft
        Case IIf(vbwProfiler.vbwExecuteLine(6451), VBWPROFILER_EMPTY, _
        Hovercraft)
vbwProfiler.vbwExecuteLine 6452
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6453
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6454
           AddPCLproperty "Index", .index, wdNumber, "Index"
vbwProfiler.vbwExecuteLine 6455
           AddPCLproperty "Hovercraft Type", .subtype, wdList, "SubType", "GEV skirt", "SEV sidewalls"
            'note, no empty space allowed
vbwProfiler.vbwExecuteLine 6456
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
vbwProfiler.vbwExecuteLine 6457
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
vbwProfiler.vbwExecuteLine 6458
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
vbwProfiler.vbwExecuteLine 6459
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
vbwProfiler.vbwExecuteLine 6460
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
vbwProfiler.vbwExecuteLine 6461
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
vbwProfiler.vbwExecuteLine 6462
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6463
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
vbwProfiler.vbwExecuteLine 6464
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6465
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6466
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6467
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6468
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6469:        Case Superstructure
        Case IIf(vbwProfiler.vbwExecuteLine(6469), VBWPROFILER_EMPTY, _
        Superstructure)
vbwProfiler.vbwExecuteLine 6470
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6471
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6472
           AddPCLproperty "Index", .index, wdNumber, "Index"
vbwProfiler.vbwExecuteLine 6473
           AddPCLproperty "Orientation", .Orientation, wdList, "Orientation", "top", "underside", "front", "back", "left", "right"
vbwProfiler.vbwExecuteLine 6474
           AddPCLproperty "Compartmentalization", .Compartmentalization, wdList, "Compartmentalization", "none", "heavy", "total"
vbwProfiler.vbwExecuteLine 6475
           AddPCLproperty "Top Deck", .TopDeck, wdBool, "TopDeck"
vbwProfiler.vbwExecuteLine 6476
           AddPCLproperty "% Covered Deck", .PercentCovered, wdNumber, "PercentCovered"
vbwProfiler.vbwExecuteLine 6477
           AddPCLproperty "% Flight Deck", .PercentFlightDeck, wdNumber, "PercentFlightDeck"
vbwProfiler.vbwExecuteLine 6478
           AddPCLproperty "Flight Deck Option", .flightdeckoption, wdList, "FlightDeckOption", "none", "landing pad", "angled flight deck"
vbwProfiler.vbwExecuteLine 6479
           AddPCLproperty "Slope Right", .SlopeR, wdList, "sloper", "none", "30 degrees", "60 degrees"
vbwProfiler.vbwExecuteLine 6480
           AddPCLproperty "Slope Left", .slopel, wdList, "slopel", "none", "30 degrees", "60 degrees"
vbwProfiler.vbwExecuteLine 6481
           AddPCLproperty "Slope Front", .slopef, wdList, "slopeF", "none", "30 degrees", "60 degrees"
vbwProfiler.vbwExecuteLine 6482
           AddPCLproperty "Slope Back", .slopeb, wdList, "slopeb", "none", "30 degrees", "60 degrees"
vbwProfiler.vbwExecuteLine 6483
           AddPCLproperty "Empty Space", .EmptySpace, wdDouble, "EmptySpace"
vbwProfiler.vbwExecuteLine 6484
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
vbwProfiler.vbwExecuteLine 6485
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
vbwProfiler.vbwExecuteLine 6486
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
vbwProfiler.vbwExecuteLine 6487
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
vbwProfiler.vbwExecuteLine 6488
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
vbwProfiler.vbwExecuteLine 6489
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
vbwProfiler.vbwExecuteLine 6490
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6491
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
vbwProfiler.vbwExecuteLine 6492
           AddPCLproperty "Top Deck Area", Format(.TotalDeckArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6493
           AddPCLproperty "Flight Deck Length", Format(.flightdecklength, "standard") & " ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6494
           AddPCLproperty "Flight Deck Area", Format(.FlightDeckArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6495
           AddPCLproperty "Covered Deck Area", Format(.covereddeckarea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6496
           AddPCLproperty "Deck Cost", "$" & Format(.DeckCost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6497
           AddPCLproperty "Deck Weight", Format(.DeckWeight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6498
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6499
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6500
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6501
           AddPCLproperty "Access Space", Format(.AccessSpace, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6502
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6503
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6504:        Case OpenMount
        Case IIf(vbwProfiler.vbwExecuteLine(6504), VBWPROFILER_EMPTY, _
        OpenMount)
vbwProfiler.vbwExecuteLine 6505
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6506
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6507
           AddPCLproperty "Index", .index, wdNumber, "Index"
vbwProfiler.vbwExecuteLine 6508
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6509
           AddPCLproperty "Orientation", .Orientation, wdList, "Orientation", "top", "underside", "front", "back", "left", "right"
vbwProfiler.vbwExecuteLine 6510
           AddPCLproperty "Rotation Type", .Rotation, wdList, "Rotation", "none", "full", "limited"
vbwProfiler.vbwExecuteLine 6511
           AddPCLproperty "Empty Space", .EmptySpace, wdDouble, "EmptySpace"
vbwProfiler.vbwExecuteLine 6512
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6513
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
vbwProfiler.vbwExecuteLine 6514
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6515
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6516
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6517
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6518
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6519:        Case Mast
        Case IIf(vbwProfiler.vbwExecuteLine(6519), VBWPROFILER_EMPTY, _
        Mast)
vbwProfiler.vbwExecuteLine 6520
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6521
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6522
           AddPCLproperty "Index", .index, wdNumber, "Index"
vbwProfiler.vbwExecuteLine 6523
           AddPCLproperty "# of Masts", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6524
           AddPCLproperty "Height", .Height, wdNumber, "Height"
vbwProfiler.vbwExecuteLine 6525
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "wood", "metal"
            'note no empty space allowed
vbwProfiler.vbwExecuteLine 6526
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6527
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
vbwProfiler.vbwExecuteLine 6528
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6529
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6530
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6531
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6532
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6533:        Case Pod
        Case IIf(vbwProfiler.vbwExecuteLine(6533), VBWPROFILER_EMPTY, _
        Pod)
vbwProfiler.vbwExecuteLine 6534
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6535
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6536
           AddPCLproperty "Index", .index, wdNumber, "Index"
vbwProfiler.vbwExecuteLine 6537
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6538
           AddPCLproperty "Orientation", .Orientation, wdList, "Orientation", "top", "underside", "front", "back", "left", "right"
vbwProfiler.vbwExecuteLine 6539
           AddPCLproperty "Empty Space", .EmptySpace, wdDouble, "EmptySpace"
vbwProfiler.vbwExecuteLine 6540
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
vbwProfiler.vbwExecuteLine 6541
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
vbwProfiler.vbwExecuteLine 6542
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
vbwProfiler.vbwExecuteLine 6543
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
vbwProfiler.vbwExecuteLine 6544
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
vbwProfiler.vbwExecuteLine 6545
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
vbwProfiler.vbwExecuteLine 6546
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6547
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
vbwProfiler.vbwExecuteLine 6548
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6549
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6550
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6551
           AddPCLproperty "Access Space", Format(.AccessSpace, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6552
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6553
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6554:        Case Turret, Popturret
        Case IIf(vbwProfiler.vbwExecuteLine(6554), VBWPROFILER_EMPTY, _
        Turret), Popturret
vbwProfiler.vbwExecuteLine 6555
           AddPCLproperty "Settings", "", wdText, "Disabled"
            'note: only turrets and popturrets will have a "rotation space" statistic
vbwProfiler.vbwExecuteLine 6556
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6557
           AddPCLproperty "Index", .index, wdNumber, "Index"
vbwProfiler.vbwExecuteLine 6558
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6559
           AddPCLproperty "Orientation", .Orientation, wdList, "Orientation", "top", "underside", "front", "back", "left", "right"
vbwProfiler.vbwExecuteLine 6560
           AddPCLproperty "Rotation", .Rotation, wdList, "Rotation", "none", "limited", "full"
vbwProfiler.vbwExecuteLine 6561
           AddPCLproperty "Compartmentalization", .Compartmentalization, wdList, "Compartmentalization", "none", "heavy", "total"
vbwProfiler.vbwExecuteLine 6562
           AddPCLproperty "Slope Right", .SlopeR, wdList, "sloper", "none", "30 degrees", "60 degrees"
vbwProfiler.vbwExecuteLine 6563
           AddPCLproperty "Slope Left", .slopel, wdList, "slopel", "none", "30 degrees", "60 degrees"
vbwProfiler.vbwExecuteLine 6564
           AddPCLproperty "Slope Front", .slopef, wdList, "slopeF", "none", "30 degrees", "60 degrees"
vbwProfiler.vbwExecuteLine 6565
           AddPCLproperty "Slope Back", .slopeb, wdList, "slopeb", "none", "30 degrees", "60 degrees"
vbwProfiler.vbwExecuteLine 6566
           AddPCLproperty "Empty Space", .EmptySpace, wdDouble, "EmptySpace"
vbwProfiler.vbwExecuteLine 6567
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
vbwProfiler.vbwExecuteLine 6568
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
vbwProfiler.vbwExecuteLine 6569
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
vbwProfiler.vbwExecuteLine 6570
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
vbwProfiler.vbwExecuteLine 6571
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
vbwProfiler.vbwExecuteLine 6572
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
vbwProfiler.vbwExecuteLine 6573
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6574
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
vbwProfiler.vbwExecuteLine 6575
           AddPCLproperty "Rotation Space", .RotationSpace, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6576
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6577
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6578
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6579
           AddPCLproperty "Access Space", Format(.AccessSpace, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6580
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6581
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6582:        Case Gasbag
        Case IIf(vbwProfiler.vbwExecuteLine(6582), VBWPROFILER_EMPTY, _
        Gasbag)
vbwProfiler.vbwExecuteLine 6583
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6584
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6585
           AddPCLproperty "Index", .index, wdNumber, "Index"
            'note no empty space allowed
vbwProfiler.vbwExecuteLine 6586
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6587
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
vbwProfiler.vbwExecuteLine 6588
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6589
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6590
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6591
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6592
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6593:         Case Cargo
         Case IIf(vbwProfiler.vbwExecuteLine(6593), VBWPROFILER_EMPTY, _
        Cargo)
vbwProfiler.vbwExecuteLine 6594
           AddPCLproperty "Settings", "", wdText, "Disabled"
           'AddPCLproperty "Index", .Index, wdNumber, "Index"
vbwProfiler.vbwExecuteLine 6595
           AddPCLproperty "Cargo Type", .subtype, wdList, "Subtype", "standard", "hidden", "open"
vbwProfiler.vbwExecuteLine 6596
           AddPCLproperty "Cargo Room", .CargoSpace, wdDouble, "CargoSpace"
vbwProfiler.vbwExecuteLine 6597
           AddPCLproperty "Empty Weight", .Weight, wdDouble, "Weight"
vbwProfiler.vbwExecuteLine 6598
           AddPCLproperty "Weight Per cf", .WeightPerCubicFoot, wdDouble, "WeightPerCubicFoot"
            'note no empty space allowed
vbwProfiler.vbwExecuteLine 6599
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           'AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
vbwProfiler.vbwExecuteLine 6600
           AddPCLproperty "Compartment Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6601
           AddPCLproperty "Cargo Weight", Format(.CargoWeight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6602
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6603
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"


'vbwLine 6604:        Case equipmentPod
        Case IIf(vbwProfiler.vbwExecuteLine(6604), VBWPROFILER_EMPTY, _
        equipmentPod)
vbwProfiler.vbwExecuteLine 6605
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6606
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6607
           AddPCLproperty "Index", .index, wdNumber, "Index"
vbwProfiler.vbwExecuteLine 6608
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6609
           AddPCLproperty "Empty Space", .EmptySpace, wdDouble, "EmptySpace"
vbwProfiler.vbwExecuteLine 6610
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
vbwProfiler.vbwExecuteLine 6611
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
vbwProfiler.vbwExecuteLine 6612
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
vbwProfiler.vbwExecuteLine 6613
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
vbwProfiler.vbwExecuteLine 6614
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
vbwProfiler.vbwExecuteLine 6615
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
vbwProfiler.vbwExecuteLine 6616
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6617
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
vbwProfiler.vbwExecuteLine 6618
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6619
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6620
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6621
           AddPCLproperty "Access Space", Format(.AccessSpace, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6622
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6623
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6624:       Case SideCar
       Case IIf(vbwProfiler.vbwExecuteLine(6624), VBWPROFILER_EMPTY, _
        SideCar)
vbwProfiler.vbwExecuteLine 6625
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6626
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6627
           AddPCLproperty "Index", .index, wdNumber, "Index"
            'note, no empty space allowed
vbwProfiler.vbwExecuteLine 6628
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
vbwProfiler.vbwExecuteLine 6629
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
vbwProfiler.vbwExecuteLine 6630
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
vbwProfiler.vbwExecuteLine 6631
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
vbwProfiler.vbwExecuteLine 6632
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
vbwProfiler.vbwExecuteLine 6633
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
vbwProfiler.vbwExecuteLine 6634
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6635
           AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
vbwProfiler.vbwExecuteLine 6636
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6637
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6638
           AddPCLproperty "Access Space", Format(.AccessSpace, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6639
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6640
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6641
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6642:       Case SolarPanel
       Case IIf(vbwProfiler.vbwExecuteLine(6642), VBWPROFILER_EMPTY, _
        SolarPanel)
vbwProfiler.vbwExecuteLine 6643
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6644
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
           'AddPCLproperty "Index", .Index, wdNumber, "Index"
vbwProfiler.vbwExecuteLine 6645
           AddPCLproperty "Surface Area", .SurfaceArea, wdDouble, "SurfaceArea"
vbwProfiler.vbwExecuteLine 6646
           AddPCLproperty "Retractable?", .Retractable, wdBool, "Retractable"
vbwProfiler.vbwExecuteLine 6647
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6648
           AddPCLproperty "Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy"
vbwProfiler.vbwExecuteLine 6649
           AddPCLproperty "Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced"
vbwProfiler.vbwExecuteLine 6650
           AddPCLproperty "Responsive", .Responsive, wdBool, "Responsive"
vbwProfiler.vbwExecuteLine 6651
           AddPCLproperty "Robotic", .Robotic, wdBool, "Robotic"
vbwProfiler.vbwExecuteLine 6652
           AddPCLproperty "Biomechanical", .Biomechanical, wdBool, "Biomechanical"
vbwProfiler.vbwExecuteLine 6653
           AddPCLproperty "Living Metal", .LivingMetal, wdBool, "LivingMetal"
vbwProfiler.vbwExecuteLine 6654
           AddPCLproperty "Statistics", "", wdText, "Disabled"
           'AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
vbwProfiler.vbwExecuteLine 6655
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6656
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6657
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6658
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6659
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6660:      Case SolarCellArray
      Case IIf(vbwProfiler.vbwExecuteLine(6660), VBWPROFILER_EMPTY, _
        SolarCellArray)
vbwProfiler.vbwExecuteLine 6661
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6662
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6663
           AddPCLproperty "Percent Area Covered", .PercentCovered, wdNumber, "PercentCovered"
vbwProfiler.vbwExecuteLine 6664
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6665
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6666
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6667
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kw", wdText, "Disabled"
           'AddPCLproperty "Endurance", .Endurance & " yrs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6668
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6669
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
    End Select
vbwProfiler.vbwExecuteLine 6670 'B
vbwProfiler.vbwExecuteLine 6671
End With
vbwProfiler.vbwProcOut 275
vbwProfiler.vbwExecuteLine 6672
End Sub
         
Private Sub ShowPropsForPropulsion(ByVal component As Integer, ByVal Key As String)
'////////////////////////////////////////////
'Propulsion Systems
'////////////////////////////////////////////
vbwProfiler.vbwProcIn 276
vbwProfiler.vbwExecuteLine 6673
With m_oCurrentVeh.Components(Key)

' Fill the window with properties for the correct Collection item
vbwProfiler.vbwExecuteLine 6674
Select Case component
'vbwLine 6675:        Case WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, TrackedDrivetrain, LegDrivetrain, FlexibodyDrivetrain
        Case IIf(vbwProfiler.vbwExecuteLine(6675), VBWPROFILER_EMPTY, _
        WheeledDrivetrain), AllWheelDriveWheeledDrivetrain, TrackedDrivetrain, LegDrivetrain, FlexibodyDrivetrain
vbwProfiler.vbwExecuteLine 6676
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6677
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6678
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 6679
           If component = LegDrivetrain Then
vbwProfiler.vbwExecuteLine 6680
                AddPCLproperty "Motive Power (per motor)", .motivepower, wdDouble, "MotivePower"
           Else
vbwProfiler.vbwExecuteLine 6681 'B
vbwProfiler.vbwExecuteLine 6682
                AddPCLproperty "Motive Power", .motivepower, wdDouble, "MotivePower"
           End If
vbwProfiler.vbwExecuteLine 6683 'B
vbwProfiler.vbwExecuteLine 6684
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 6685
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6686
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6687
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6688
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6689
            If component = LegDrivetrain Then
vbwProfiler.vbwExecuteLine 6690
               AddPCLproperty "Volume per leg:", Format(.Volume, "standard"), wdText, "Disabled"
            Else
vbwProfiler.vbwExecuteLine 6691 'B
vbwProfiler.vbwExecuteLine 6692
               AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
            End If
vbwProfiler.vbwExecuteLine 6693 'B
vbwProfiler.vbwExecuteLine 6694
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6695
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6696:        Case OrnithopterDrivetrain, TTRRotorDrivetrain, CARRotorDrivetrain
        Case IIf(vbwProfiler.vbwExecuteLine(6696), VBWPROFILER_EMPTY, _
        OrnithopterDrivetrain), TTRRotorDrivetrain, CARRotorDrivetrain
vbwProfiler.vbwExecuteLine 6697
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6698
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6699
           If component = OrnithopterDrivetrain Then
vbwProfiler.vbwExecuteLine 6700
                AddPCLproperty "Motive Power (per motor)", .motivepower, wdDouble, "MotivePower"
           Else
vbwProfiler.vbwExecuteLine 6701 'B
vbwProfiler.vbwExecuteLine 6702
                AddPCLproperty "Motive Power", .motivepower, wdDouble, "MotivePower"
           End If
vbwProfiler.vbwExecuteLine 6703 'B
vbwProfiler.vbwExecuteLine 6704
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 6705
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 6706
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6707
           AddPCLproperty "Lift", Format(.Lift, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6708
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6709
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6710
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6711
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6712
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6713
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6714
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6715:        Case MMRRotorDrivetrain
        Case IIf(vbwProfiler.vbwExecuteLine(6715), VBWPROFILER_EMPTY, _
        MMRRotorDrivetrain)
vbwProfiler.vbwExecuteLine 6716
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6717
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6718
           AddPCLproperty "Tilt Rotor?", .TiltRotor, wdBool, "TiltRotor"
vbwProfiler.vbwExecuteLine 6719
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 6720
           AddPCLproperty "Motive Power", .motivepower, wdDouble, "MotivePower"
vbwProfiler.vbwExecuteLine 6721
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 6722
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6723
           AddPCLproperty "Lift", Format(.Lift, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6724
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6725
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6726
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6727
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6728
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6729
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6730
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6731:        Case AerialPropeller
        Case IIf(vbwProfiler.vbwExecuteLine(6731), VBWPROFILER_EMPTY, _
        AerialPropeller)
vbwProfiler.vbwExecuteLine 6732
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6733
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6734
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6735
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 6736
           AddPCLproperty "Motive Power", .motivepower, wdDouble, "MotivePower"
vbwProfiler.vbwExecuteLine 6737
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 6738
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6739
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6740
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6741
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6742
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           'AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           'AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           'AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6743:        Case DuctedFan
        Case IIf(vbwProfiler.vbwExecuteLine(6743), VBWPROFILER_EMPTY, _
        DuctedFan)
vbwProfiler.vbwExecuteLine 6744
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6745
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6746
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6747
           AddPCLproperty "Motive Power", .motivepower, wdDouble, "MotivePower"
vbwProfiler.vbwExecuteLine 6748
           AddPCLproperty "Hover Fan", .HoverFan, wdBool, "HoverFan"
vbwProfiler.vbwExecuteLine 6749
           AddPCLproperty "Lift Engine", .LiftEngine, wdBool, "LiftEngine"
vbwProfiler.vbwExecuteLine 6750
           AddPCLproperty "Vectored Thrust", .VectoredThrust, wdBool, "VectoredThrust"
vbwProfiler.vbwExecuteLine 6751
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 6752
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 6753
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6754
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6755
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6756
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6757
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6758
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6759
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6760
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"



'vbwLine 6761:        Case PaddleWheel, ScrewPropeller, lightScrewPropeller, DuctedPropeller, Hydrojet, MHDTunnel
        Case IIf(vbwProfiler.vbwExecuteLine(6761), VBWPROFILER_EMPTY, _
        PaddleWheel), ScrewPropeller, lightScrewPropeller, DuctedPropeller, Hydrojet, MHDTunnel
vbwProfiler.vbwExecuteLine 6762
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6763
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6764
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6765
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 6766
           AddPCLproperty "Motive Power", .motivepower, wdDouble, "MotivePower"
vbwProfiler.vbwExecuteLine 6767
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 6768
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6769
           AddPCLproperty "Aquatic Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6770
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6771
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6772
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6773
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6774
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6775
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6776:        Case RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness
        Case IIf(vbwProfiler.vbwExecuteLine(6776), VBWPROFILER_EMPTY, _
        RopeHarness), YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness
vbwProfiler.vbwExecuteLine 6777
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6778
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6779
           AddPCLproperty "Animal Type", .subtype, wdList, "SubType", "Land Animal", "Swimming Animal", "Flying Animal"
vbwProfiler.vbwExecuteLine 6780
           AddPCLproperty "Animal Description", .AnimalDescription, wdText, "AnimalDescription"
vbwProfiler.vbwExecuteLine 6781
           AddPCLproperty "Strength per Animal", .BeastST, wdNumber, "BeastST"
vbwProfiler.vbwExecuteLine 6782
           AddPCLproperty "Hexes Per Animal", .Hexes, wdNumber, "Hexes"
vbwProfiler.vbwExecuteLine 6783
           AddPCLproperty "Number of Animals", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6784
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6785
            If .subtype = "Land Animal" Then
vbwProfiler.vbwExecuteLine 6786
               AddPCLproperty "Motive Power", .motivepower, wdText, "Disabled"
            Else
vbwProfiler.vbwExecuteLine 6787 'B
vbwProfiler.vbwExecuteLine 6788
               AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
            End If
vbwProfiler.vbwExecuteLine 6789 'B
vbwProfiler.vbwExecuteLine 6790
           AddPCLproperty "Total Hexes of Animals", .TotalHexes, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6791
           AddPCLproperty "Move per Animal", .Move, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6792
           AddPCLproperty "Speed per Animal", .Speed, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6793
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6794
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"



'vbwLine 6795:        Case RowingPositions
        Case IIf(vbwProfiler.vbwExecuteLine(6795), VBWPROFILER_EMPTY, _
        RowingPositions)
vbwProfiler.vbwExecuteLine 6796
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6797
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6798
           AddPCLproperty "# of Positions", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6799
           AddPCLproperty "Avg. ST per Position", .RowerST, wdNumber, "RowerST"
vbwProfiler.vbwExecuteLine 6800
           AddPCLproperty "DR per Position", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 6801
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 6802
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6803
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6804
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6805
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6806
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6807
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6808
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6809:        Case FullRig, SquareRig, ForeandAftRig, AerialSail, AerialSailForeAftRig
        Case IIf(vbwProfiler.vbwExecuteLine(6809), VBWPROFILER_EMPTY, _
        FullRig), SquareRig, ForeandAftRig, AerialSail, AerialSailForeAftRig
vbwProfiler.vbwExecuteLine 6810
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6811
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6812
           AddPCLproperty "Sail Material", .material, wdList, "Material", "cloth", "synthetic", "bioplas"
vbwProfiler.vbwExecuteLine 6813
           AddPCLproperty "Wind", .Wind, wdList, "Wind", "calm", "light air", "light breeze", "gentle breeze", "moderate breeze", "fresh breeze", "strong breeze", "gale force winds"
vbwProfiler.vbwExecuteLine 6814
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 6815
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6816
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6817
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6818
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6819
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6820
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6821:        Case lightSail
        Case IIf(vbwProfiler.vbwExecuteLine(6821), VBWPROFILER_EMPTY, _
        lightSail)
vbwProfiler.vbwExecuteLine 6822
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6823
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6824
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 6825
           AddPCLproperty "Sail Size (sq mi)", .SurfaceArea, wdDouble, "SurfaceArea"
vbwProfiler.vbwExecuteLine 6826
           AddPCLproperty "AU Distance", .AUDistance, wdDouble, "AUDistance"
vbwProfiler.vbwExecuteLine 6827
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 6828
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6829
           AddPCLproperty "Motive Thrust", Format(.Thrust, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6830
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6831
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6832
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6833
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6834:        Case Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan
        Case IIf(vbwProfiler.vbwExecuteLine(6834), VBWPROFILER_EMPTY, _
        Turbojet), Turbofan, Ramjet, TurboRamjet, Hyperfan
vbwProfiler.vbwExecuteLine 6835
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6836
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6837
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6838
           AddPCLproperty "Thrust (lbs)", .DesiredThrust, wdDouble, "DesiredThrust"
vbwProfiler.vbwExecuteLine 6839
           AddPCLproperty "Afterburner", .Afterburner, wdBool, "Afterburner"
vbwProfiler.vbwExecuteLine 6840
           If component <> Ramjet Then '//ramjets cant be lift engines because they need air travelling through them in forward motion
vbwProfiler.vbwExecuteLine 6841
                AddPCLproperty "Lift Engine", .LiftEngine, wdBool, "LiftEngine"
           End If
vbwProfiler.vbwExecuteLine 6842 'B
vbwProfiler.vbwExecuteLine 6843
           AddPCLproperty "Vectored Thrust", .VectoredThrust, wdBool, "VectoredThrust"
vbwProfiler.vbwExecuteLine 6844
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 6845
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 6846
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6847
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6848
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6849
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6850
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6851
           AddPCLproperty "Fuel Consumption", Format(.FuelConsumption, "standard") & " gph", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6852
           AddPCLproperty "AB Thrust", Format(.ABThrust, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6853
           AddPCLproperty "AB Lit Fuel Consumption", Format(.ABConsumption, "standard") & " gph", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6854
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6855
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6856
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6857
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6858
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6859:        Case FusionAirRam 'only jet engine that cant use afterburner
        Case IIf(vbwProfiler.vbwExecuteLine(6859), VBWPROFILER_EMPTY, _
        FusionAirRam) 'only jet engine that cant use afterburner
vbwProfiler.vbwExecuteLine 6860
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6861
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6862
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6863
           AddPCLproperty "Thrust (lbs)", .DesiredThrust, wdDouble, "DesiredThrust"
vbwProfiler.vbwExecuteLine 6864
           AddPCLproperty "Lift Engine", .LiftEngine, wdBool, "LiftEngine"
vbwProfiler.vbwExecuteLine 6865
           AddPCLproperty "Vectored Thrust", .VectoredThrust, wdBool, "VectoredThrust"
vbwProfiler.vbwExecuteLine 6866
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 6867
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 6868
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6869
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6870
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6871
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"

vbwProfiler.vbwExecuteLine 6872
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6873
           AddPCLproperty "Endurance", Format(.FuelConsumption, "standard") & " yrs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6874
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6875
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6876
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6877
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6878
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6879:        Case StandardThruster, SuperThruster, MegaThruster
        Case IIf(vbwProfiler.vbwExecuteLine(6879), VBWPROFILER_EMPTY, _
        StandardThruster), SuperThruster, MegaThruster
vbwProfiler.vbwExecuteLine 6880
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6881
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6882
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6883
           AddPCLproperty "Thrust (lbs)", .DesiredThrust, wdDouble, "DesiredThrust"
vbwProfiler.vbwExecuteLine 6884
           AddPCLproperty "Lift Engine", .LiftEngine, wdBool, "LiftEngine"
vbwProfiler.vbwExecuteLine 6885
           AddPCLproperty "Vectored Thrust", .VectoredThrust, wdBool, "VectoredThrust"
vbwProfiler.vbwExecuteLine 6886
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 6887
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 6888
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6889
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6890
           AddPCLproperty "Power", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6891
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6892
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6893
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6894
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6895
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6896:        Case LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal
        Case IIf(vbwProfiler.vbwExecuteLine(6896), VBWPROFILER_EMPTY, _
        LiquidFuelRocket), MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal
vbwProfiler.vbwExecuteLine 6897
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6898
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6899
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6900
           AddPCLproperty "Thrust (lbs)", .DesiredThrust, wdDouble, "DesiredThrust"
vbwProfiler.vbwExecuteLine 6901
           AddPCLproperty "Lift Engine", .LiftEngine, wdBool, "LiftEngine"
vbwProfiler.vbwExecuteLine 6902
           AddPCLproperty "Vectored Thrust", .VectoredThrust, wdBool, "VectoredThrust"
vbwProfiler.vbwExecuteLine 6903
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 6904
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 6905
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6906
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6907
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6908
           AddPCLproperty "Fuel Consumption", Format(.FuelConsumption, "standard") & " gph", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6909
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6910
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6911
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6912
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6913
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6914:        Case AntimatterPion
        Case IIf(vbwProfiler.vbwExecuteLine(6914), VBWPROFILER_EMPTY, _
        AntimatterPion)
vbwProfiler.vbwExecuteLine 6915
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6916
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6917
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6918
           AddPCLproperty "Thrust (lbs)", .DesiredThrust, wdDouble, "DesiredThrust"
vbwProfiler.vbwExecuteLine 6919
           AddPCLproperty "Lift Engine", .LiftEngine, wdBool, "LiftEngine"
vbwProfiler.vbwExecuteLine 6920
           AddPCLproperty "Vectored Thrust", .VectoredThrust, wdBool, "VectoredThrust"
vbwProfiler.vbwExecuteLine 6921
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 6922
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 6923
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6924
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6925
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6926
           AddPCLproperty "Antimatter Fuel Consumption", Format(.FuelConsumption, "standard") & " grams per hour", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6927
           AddPCLproperty "Hydrogen Fuel Consumption", Format(.FuelConsumption2, "standard") & " gph", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6928
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6929
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6930
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6931
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6932
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6933:        Case SolidRocketEngine
        Case IIf(vbwProfiler.vbwExecuteLine(6933), VBWPROFILER_EMPTY, _
        SolidRocketEngine)
vbwProfiler.vbwExecuteLine 6934
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6935
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6936
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6937
           AddPCLproperty "Thrust (lbs)", .DesiredThrust, wdDouble, "DesiredThrust"
vbwProfiler.vbwExecuteLine 6938
           AddPCLproperty "Burn Time (mins)", .BurnTime, wdDouble, "BurnTime"
vbwProfiler.vbwExecuteLine 6939
           AddPCLproperty "Lift Engine", .LiftEngine, wdBool, "LiftEngine"
vbwProfiler.vbwExecuteLine 6940
           AddPCLproperty "Vectored Thrust", .VectoredThrust, wdBool, "VectoredThrust"
vbwProfiler.vbwExecuteLine 6941
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 6942
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 6943
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6944
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6945
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6946
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6947
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6948
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6949
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6950:        Case OrionEngine
        Case IIf(vbwProfiler.vbwExecuteLine(6950), VBWPROFILER_EMPTY, _
        OrionEngine)
vbwProfiler.vbwExecuteLine 6951
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6952
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6953
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6954
           AddPCLproperty "Pulse Rate (bps)", .PulseRate, wdDouble, "PulseRate"
vbwProfiler.vbwExecuteLine 6955
           AddPCLproperty "Bomb Size (kt)", .BombSize, wdDouble, "BombSize"
vbwProfiler.vbwExecuteLine 6956
           AddPCLproperty "# of Bombs", .NumBombs, wdNumber, "NumBombs"
vbwProfiler.vbwExecuteLine 6957
           AddPCLproperty "Lift Engine", .LiftEngine, wdBool, "LiftEngine"
vbwProfiler.vbwExecuteLine 6958
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 6959
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 6960
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6961
           AddPCLproperty "Motive Thrust", Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6962
           AddPCLproperty "Thrust Time (secs)", .ThrustTime, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6963
           AddPCLproperty "Bomb Weight", .BombWeight, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6964
           AddPCLproperty "Bomb Cost", .BombCost, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6965
           AddPCLproperty "Bomb Volume", .BombVolume, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6966
           AddPCLproperty "Total Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6967
           AddPCLproperty "Total Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6968
           AddPCLproperty "Total Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6969
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6970
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6971:        Case MagLevLifter
        Case IIf(vbwProfiler.vbwExecuteLine(6971), VBWPROFILER_EMPTY, _
        MagLevLifter)
vbwProfiler.vbwExecuteLine 6972
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6973
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6974
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6975
           AddPCLproperty "Lift (lbs)", .DesiredLift, wdDouble, "DesiredLift"
vbwProfiler.vbwExecuteLine 6976
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 6977
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 6978
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6979
           AddPCLproperty "Total lift", Format(.Lift, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6980
           AddPCLproperty "Power", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6981
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6982
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6983
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6984
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6985
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 6986:        Case JumpDrive, TeleportationDrive
        Case IIf(vbwProfiler.vbwExecuteLine(6986), VBWPROFILER_EMPTY, _
        JumpDrive), TeleportationDrive
vbwProfiler.vbwExecuteLine 6987
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6988
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 6989
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 6990
           AddPCLproperty "Capacity (tons)", .DesiredCapacity, wdDouble, "DesiredCapacity"
vbwProfiler.vbwExecuteLine 6991
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 6992
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 6993
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6994
           AddPCLproperty "Total Capacity", Format(.capacity, "standard") & " tons", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6995
           AddPCLproperty "Power", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6996
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6997
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6998
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 6999
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7000
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7001:        Case Hyperdrive
        Case IIf(vbwProfiler.vbwExecuteLine(7001), VBWPROFILER_EMPTY, _
        Hyperdrive)
vbwProfiler.vbwExecuteLine 7002
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7003
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7004
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7005
           AddPCLproperty "Capacity (tons)", .DesiredCapacity, wdDouble, "DesiredCapacity"
vbwProfiler.vbwExecuteLine 7006
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7007
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7008
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7009
           AddPCLproperty "Total Capacity", Format(.capacity, "standard") & " tons", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7010
           AddPCLproperty "Initial Power", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7011
           AddPCLproperty "Sustained Power", Format(.SustainedPower, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7012
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7013
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7014
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7015
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7016
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7017:        Case WarpDrive
        Case IIf(vbwProfiler.vbwExecuteLine(7017), VBWPROFILER_EMPTY, _
        WarpDrive)
vbwProfiler.vbwExecuteLine 7018
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7019
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7020
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7021
           AddPCLproperty "Warp Thrust Factor", .DesiredCapacity, wdDouble, "DesiredCapacity"
vbwProfiler.vbwExecuteLine 7022
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7023
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7024
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7025
           AddPCLproperty "Total WTF", Format(.capacity, "standard") & " WTF", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7026
           AddPCLproperty "Power", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7027
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7028
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7029
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7030
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7031
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7032:        Case SubQuantumConveyor, QuantumConveyor, TwoQuantumConveyor
        Case IIf(vbwProfiler.vbwExecuteLine(7032), VBWPROFILER_EMPTY, _
        SubQuantumConveyor), QuantumConveyor, TwoQuantumConveyor
vbwProfiler.vbwExecuteLine 7033
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7034
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7035
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7036
           AddPCLproperty "Max Transport Weight (lbs)", .DesiredCapacity, wdDouble, "desiredCapacity"
vbwProfiler.vbwExecuteLine 7037
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7038
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7039
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7040
           AddPCLproperty "Total Transport Weight", Format(.capacity, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7041
           AddPCLproperty "Power", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7042
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7043
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7044
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7045
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7046
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


    '/////////////////////////////////////////
    ' Aerostatic Lift Systems
'vbwLine 7047:        Case HotAir, Hydrogen, Helium
        Case IIf(vbwProfiler.vbwExecuteLine(7047), VBWPROFILER_EMPTY, _
        HotAir), Hydrogen, Helium
vbwProfiler.vbwExecuteLine 7048
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7049
           AddPCLproperty "Useful Static Lift (lbs)", .Lift, wdDouble, "Lift"
vbwProfiler.vbwExecuteLine 7050
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7051
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7052
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"


'vbwLine 7053:        Case ContraGravGenerator
        Case IIf(vbwProfiler.vbwExecuteLine(7053), VBWPROFILER_EMPTY, _
        ContraGravGenerator)
vbwProfiler.vbwExecuteLine 7054
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7055
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7056
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7057
           AddPCLproperty "Lift", .DesiredLift, wdDouble, "DesiredLift"
vbwProfiler.vbwExecuteLine 7058
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7059
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7060
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7061
           AddPCLproperty "Total Lift", Format(.Lift, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7062
           AddPCLproperty "Power", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7063
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7064
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7065
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7066
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7067
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

End Select
vbwProfiler.vbwExecuteLine 7068 'B
vbwProfiler.vbwExecuteLine 7069
End With
vbwProfiler.vbwProcOut 276
vbwProfiler.vbwExecuteLine 7070
End Sub

Private Sub ShowPropsforInstruments(ByVal component As Integer, ByVal Key As String)
'///////////////////////////////////////////
'Instruments and Electronics
vbwProfiler.vbwProcIn 277
vbwProfiler.vbwExecuteLine 7071
With m_oCurrentVeh.Components(Key)

vbwProfiler.vbwExecuteLine 7072
Select Case component
' Fill the window with properties for the correct Collection item
'vbwLine 7073:    Case RadioDirectionFinder, RadioCommunicator, TightBeamRadio, VLFRadio, CellularPhone, CellularPhonewithRadio, RadioJammer, ElfReceiver, LaserCommunicator, NeutrinoCommunicator, GravityRippleCommunicator
    Case IIf(vbwProfiler.vbwExecuteLine(7073), VBWPROFILER_EMPTY, _
        RadioDirectionFinder), RadioCommunicator, TightBeamRadio, VLFRadio, CellularPhone, CellularPhonewithRadio, RadioJammer, ElfReceiver, LaserCommunicator, NeutrinoCommunicator, GravityRippleCommunicator
vbwProfiler.vbwExecuteLine 7074
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7075
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7076
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7077
           AddPCLproperty "Desired Range", .DesiredRange, wdList, "DesiredRange", "short", "medium", "long", "very long", "extreme"
vbwProfiler.vbwExecuteLine 7078
           AddPCLproperty "Sensitivity", .Sensitivity, wdList, "Sensitivity", "normal", "sensitive", "very sensitive"
vbwProfiler.vbwExecuteLine 7079
           AddPCLproperty "FTL", .FTL, wdBool, "FTL"
vbwProfiler.vbwExecuteLine 7080
           AddPCLproperty "Receive Only", .ReceiveOnly, wdBool, "ReceiveOnly"
vbwProfiler.vbwExecuteLine 7081
           AddPCLproperty "Scrambler", .Scrambler, wdBool, "Scrambler"
vbwProfiler.vbwExecuteLine 7082
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7083
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7084
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7085
            If .FTL = False Then
vbwProfiler.vbwExecuteLine 7086
               AddPCLproperty "Actual Range", Format(.Range, "standard") & " miles", wdText, "Disabled"
            Else
vbwProfiler.vbwExecuteLine 7087 'B
vbwProfiler.vbwExecuteLine 7088
               AddPCLproperty "Actual Range", Format(.Range, "standard") & " parsecs", wdText, "Disabled"
            End If
vbwProfiler.vbwExecuteLine 7089 'B
vbwProfiler.vbwExecuteLine 7090
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7091
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7092
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7093
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7094
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7095
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7096:    Case Headlight, Searchlight, InfraredSearchlight
    Case IIf(vbwProfiler.vbwExecuteLine(7096), VBWPROFILER_EMPTY, _
        Headlight), Searchlight, InfraredSearchlight
vbwProfiler.vbwExecuteLine 7097
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7098
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7099
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7100
           AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
vbwProfiler.vbwExecuteLine 7101
           If component = Headlight Then
vbwProfiler.vbwExecuteLine 7102
               AddPCLproperty "Range (yards)", .Range, wdDouble, "Range"
           Else
vbwProfiler.vbwExecuteLine 7103 'B
vbwProfiler.vbwExecuteLine 7104
              AddPCLproperty "Range (miles)", .Range, wdDouble, "Range"
           End If
vbwProfiler.vbwExecuteLine 7105 'B
vbwProfiler.vbwExecuteLine 7106
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7107
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7108
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7109
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7110
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7111
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7112
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7113
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7114
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7115:    Case AstronomicalInstruments, Telescope, lightAmplification, LowlightTV, ExtendableSensorPeriscope
    Case IIf(vbwProfiler.vbwExecuteLine(7115), VBWPROFILER_EMPTY, _
        AstronomicalInstruments), Telescope, lightAmplification, LowlightTV, ExtendableSensorPeriscope
vbwProfiler.vbwExecuteLine 7116
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7117
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7118
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7119
           AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
vbwProfiler.vbwExecuteLine 7120
            If component = lightAmplification Then
'vbwLine 7121:            ElseIf component = ExtendableSensorPeriscope Then
            ElseIf vbwProfiler.vbwExecuteLine(7121) Or component = ExtendableSensorPeriscope Then
vbwProfiler.vbwExecuteLine 7122
               AddPCLproperty "Periscope Length", .Length, wdDouble, "Length"
            Else
vbwProfiler.vbwExecuteLine 7123 'B
vbwProfiler.vbwExecuteLine 7124
               AddPCLproperty "Magnification", .Magnification, wdDouble, "Magnification"
            End If
vbwProfiler.vbwExecuteLine 7125 'B
vbwProfiler.vbwExecuteLine 7126
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7127
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7128
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7129
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7130
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7131
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7132
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7133
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7134
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7135:    Case Radar, Ladar, NavigationalRadar, AntiCollisionRadar, AESA, LowResImagingRadar, HiResImagingRadar
    Case IIf(vbwProfiler.vbwExecuteLine(7135), VBWPROFILER_EMPTY, _
        Radar), Ladar, NavigationalRadar, AntiCollisionRadar, AESA, LowResImagingRadar, HiResImagingRadar
vbwProfiler.vbwExecuteLine 7136
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7137
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7138
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7139
           AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
vbwProfiler.vbwExecuteLine 7140
           AddPCLproperty "Range", .Range, wdDouble, "Range"
vbwProfiler.vbwExecuteLine 7141
            If component = NavigationalRadar Then
'vbwLine 7142:            ElseIf component = AntiCollisionRadar Then
            ElseIf vbwProfiler.vbwExecuteLine(7142) Or component = AntiCollisionRadar Then
            Else
vbwProfiler.vbwExecuteLine 7143 'B
vbwProfiler.vbwExecuteLine 7144
               AddPCLproperty "No Targeting", .NoTargeting, wdBool, "NoTargeting"
vbwProfiler.vbwExecuteLine 7145
               AddPCLproperty "Search Optimization", .SearchOption, wdList, "SearchOption", "none", "surface search", "air search"
vbwProfiler.vbwExecuteLine 7146
               AddPCLproperty "FTL Option", .FTL, wdBool, "FTL"
            End If
vbwProfiler.vbwExecuteLine 7147 'B
vbwProfiler.vbwExecuteLine 7148
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7149
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7150
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7151
           AddPCLproperty "Scan Rating", .ScanRating, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7152
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7153
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7154
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7155
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7156
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7157
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7158:    Case ActiveSonar, PassiveSonar
    Case IIf(vbwProfiler.vbwExecuteLine(7158), VBWPROFILER_EMPTY, _
        ActiveSonar), PassiveSonar
vbwProfiler.vbwExecuteLine 7159
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7160
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7161
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7162
            If component = ActiveSonar Then
vbwProfiler.vbwExecuteLine 7163
               AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
            End If
vbwProfiler.vbwExecuteLine 7164 'B
vbwProfiler.vbwExecuteLine 7165
           AddPCLproperty "Range", .Range, wdDouble, "Range"
vbwProfiler.vbwExecuteLine 7166
            If component = ActiveSonar Then
vbwProfiler.vbwExecuteLine 7167
               AddPCLproperty "Active / Passive?", .ActivePassive, wdBool, "ActivePassive"
vbwProfiler.vbwExecuteLine 7168
               AddPCLproperty "Depth Finding?", .DepthFinding, wdBool, "DepthFinding"
vbwProfiler.vbwExecuteLine 7169
               AddPCLproperty "Dipping Sonar?", .DippingSonar, wdBool, "DippingSonar"
vbwProfiler.vbwExecuteLine 7170
               AddPCLproperty "No Targeting?", .NoTargeting, wdBool, "NoTargeting"
            Else
vbwProfiler.vbwExecuteLine 7171 'B
vbwProfiler.vbwExecuteLine 7172
               AddPCLproperty "Dipping Sonar?", .DippingSonar, wdBool, "DippingSonar"
vbwProfiler.vbwExecuteLine 7173
               AddPCLproperty "Towed Array?", .TowedArray, wdBool, "TowedArray"
            End If
vbwProfiler.vbwExecuteLine 7174 'B
vbwProfiler.vbwExecuteLine 7175
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7176
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7177
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7178
           AddPCLproperty "Scan Rating", .ScanRating, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7179
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7180
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7181
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7182
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7183
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7184
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7185:    Case PassiveInfrared, Thermograph, PassiveRadar, PESA
    Case IIf(vbwProfiler.vbwExecuteLine(7185), VBWPROFILER_EMPTY, _
        PassiveInfrared), Thermograph, PassiveRadar, PESA
vbwProfiler.vbwExecuteLine 7186
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7187
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7188
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7189
           AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
vbwProfiler.vbwExecuteLine 7190
           AddPCLproperty "Range", .Range, wdDouble, "Range"
vbwProfiler.vbwExecuteLine 7191
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7192
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7193
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7194
           AddPCLproperty "Scan Rating", .ScanRating, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7195
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7196
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7197
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7198
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7199
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7200
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7201:    Case Geophone, MAD, MultiScanner, ChemScanner, RadScanner, BioScanner, GravScanner
    Case IIf(vbwProfiler.vbwExecuteLine(7201), VBWPROFILER_EMPTY, _
        Geophone), MAD, MultiScanner, ChemScanner, RadScanner, BioScanner, GravScanner
vbwProfiler.vbwExecuteLine 7202
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7203
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7204
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7205
           AddPCLproperty "Range", .Range, wdDouble, "Range"
vbwProfiler.vbwExecuteLine 7206
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7207
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7208
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7209
           AddPCLproperty "Scan Rating", .ScanRating, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7210
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7211
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7212
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7213
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7214
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7215
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7216:    Case RangingSoundDetector, SurveillanceSoundDetector
    Case IIf(vbwProfiler.vbwExecuteLine(7216), VBWPROFILER_EMPTY, _
        RangingSoundDetector), SurveillanceSoundDetector
vbwProfiler.vbwExecuteLine 7217
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7218
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7219
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7220
           AddPCLproperty "Sensitivity Level", .Level, wdNumber, "Level"
vbwProfiler.vbwExecuteLine 7221
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7222
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7223
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7224
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7225
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7226
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7227
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7228
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7229
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7230:    Case MeteorologicalInstruments, LowResPlanetarySurveyArray, MedResPlanetarySurveyArray, HighResPlanetarySurveyArray
    Case IIf(vbwProfiler.vbwExecuteLine(7230), VBWPROFILER_EMPTY, _
        MeteorologicalInstruments), LowResPlanetarySurveyArray, MedResPlanetarySurveyArray, HighResPlanetarySurveyArray
vbwProfiler.vbwExecuteLine 7231
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7232
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7233
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7234
           AddPCLproperty "Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside"
vbwProfiler.vbwExecuteLine 7235
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7236
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7237
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7238
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7239
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7240
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7241
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7242
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7243
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7244:    Case SoundSystem, FlightRecorder
    Case IIf(vbwProfiler.vbwExecuteLine(7244), VBWPROFILER_EMPTY, _
        SoundSystem), FlightRecorder
vbwProfiler.vbwExecuteLine 7245
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7246
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7247
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7248
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7249
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7250
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7251
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7252
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7253
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7254
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7255
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7256
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7257:    Case VehicleCamera, DigitalVehicleCamera, ReconCamera, DigitalReconCamera
    Case IIf(vbwProfiler.vbwExecuteLine(7257), VBWPROFILER_EMPTY, _
        VehicleCamera), DigitalVehicleCamera, ReconCamera, DigitalReconCamera
vbwProfiler.vbwExecuteLine 7258
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7259
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7260
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7261
           AddPCLproperty "Low light?", .Lowlight, wdBool, "Lowlight"
vbwProfiler.vbwExecuteLine 7262
           AddPCLproperty "Infrared?", .Infrared, wdBool, "Infrared"
vbwProfiler.vbwExecuteLine 7263
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7264
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7265
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7266
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7267
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7268
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7269
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7270
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7271
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7272:    Case NavigationInstruments, AutoPilot, IFF, Transponder, INS, GPS, MilitaryGPS, TFR
    Case IIf(vbwProfiler.vbwExecuteLine(7272), VBWPROFILER_EMPTY, _
        NavigationInstruments), AutoPilot, IFF, Transponder, INS, GPS, MilitaryGPS, TFR
vbwProfiler.vbwExecuteLine 7273
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7274
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7275
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7276
           If component = NavigationInstruments Then
vbwProfiler.vbwExecuteLine 7277
               AddPCLproperty "Precision?", .Precision, wdBool, "Precision"
           End If
vbwProfiler.vbwExecuteLine 7278 'B
vbwProfiler.vbwExecuteLine 7279
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7280
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7281
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7282
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7283
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7284
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7285
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7286
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7287
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7288:    Case ImprovedOpticalBombSight, AdvancedOpticalBombSight, OpticalBombSight, FireDirectionCenter, HUDWAC, PupilHUDWAC, LaserRangeFinder, LaserDesignator, LaserSpotTracker
    Case IIf(vbwProfiler.vbwExecuteLine(7288), VBWPROFILER_EMPTY, _
        ImprovedOpticalBombSight), AdvancedOpticalBombSight, OpticalBombSight, FireDirectionCenter, HUDWAC, PupilHUDWAC, LaserRangeFinder, LaserDesignator, LaserSpotTracker
vbwProfiler.vbwExecuteLine 7289
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7290
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7291
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7292
           If component = LaserDesignator Or component = LaserRangeFinder Then
vbwProfiler.vbwExecuteLine 7293
               AddPCLproperty "Range", .Range, wdDouble, "Range"
           End If
vbwProfiler.vbwExecuteLine 7294 'B
vbwProfiler.vbwExecuteLine 7295
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7296
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7297
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7298
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7299
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7300
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7301
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7302
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7303
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7304:    Case RadarDetector, LaserSensor, LaserRadarDetector, AreaRadarJammer, DeceptiveRadarJammer, InfraredJammer, ChaffDecoyDischarger, SmokeDecoyDischarger, FlareDecoyDischarger, SonarDecoyDischarger, HotSmokeDecoyDischarger, PrismDecoyDischarger, BlackOutGasDecoyDischarger, RadarReflector, BlipEnhancer, TEMPEST
    Case IIf(vbwProfiler.vbwExecuteLine(7304), VBWPROFILER_EMPTY, _
        RadarDetector), LaserSensor, LaserRadarDetector, AreaRadarJammer, DeceptiveRadarJammer, InfraredJammer, ChaffDecoyDischarger, SmokeDecoyDischarger, FlareDecoyDischarger, SonarDecoyDischarger, HotSmokeDecoyDischarger, PrismDecoyDischarger, BlackOutGasDecoyDischarger, RadarReflector, BlipEnhancer, TEMPEST
vbwProfiler.vbwExecuteLine 7305
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7306
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7307
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7308
            Select Case component
'vbwLine 7309:                Case AreaRadarJammer, DeceptiveRadarJammer, InfraredJammer
                Case IIf(vbwProfiler.vbwExecuteLine(7309), VBWPROFILER_EMPTY, _
        AreaRadarJammer), DeceptiveRadarJammer, InfraredJammer
vbwProfiler.vbwExecuteLine 7310
                    AddPCLproperty "Jammer Rating", .JammerRating, wdList, "JammerRating", 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20
'vbwLine 7311:                Case RadarDetector, LaserSensor, LaserRadarDetector
                Case IIf(vbwProfiler.vbwExecuteLine(7311), VBWPROFILER_EMPTY, _
        RadarDetector), LaserSensor, LaserRadarDetector
vbwProfiler.vbwExecuteLine 7312
                    AddPCLproperty "Advanced Version?", .ADVANCED, wdBool, "advanced"
            End Select
vbwProfiler.vbwExecuteLine 7313 'B
vbwProfiler.vbwExecuteLine 7314
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7315
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7316
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7317
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7318
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7319
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7320
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7321
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7322
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

'vbwLine 7323:    Case DecoyChaff, DecoySmoke, DecoyFlares, DecoySonarDecoy, DecoyHotSmoke, DecoyPrism, DecoyBlackOutGas
    Case IIf(vbwProfiler.vbwExecuteLine(7323), VBWPROFILER_EMPTY, _
        DecoyChaff), DecoySmoke, DecoyFlares, DecoySonarDecoy, DecoyHotSmoke, DecoyPrism, DecoyBlackOutGas
vbwProfiler.vbwExecuteLine 7324
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7325
           AddPCLproperty "Tech level", .TL, wdList, "TL", 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7326
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"

vbwProfiler.vbwExecuteLine 7327
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7328
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7329
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7330
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7331
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7332
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7333
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

'vbwLine 7334:    Case MacroFrame, MainFrame, MicroFrame, MiniComputer, SmallComputer
    Case IIf(vbwProfiler.vbwExecuteLine(7334), VBWPROFILER_EMPTY, _
        MacroFrame), MainFrame, MicroFrame, MiniComputer, SmallComputer
vbwProfiler.vbwExecuteLine 7335
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7336
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7337
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7338
           AddPCLproperty "Intelligence", .Intelligence, wdList, "Intelligence", "normal", "dumb", "genius"
vbwProfiler.vbwExecuteLine 7339
           AddPCLproperty "Configuration", .Configuration, wdList, "Configuration", "normal", "neural-net", "sentient"
vbwProfiler.vbwExecuteLine 7340
           AddPCLproperty "Compact?", .Compact, wdBool, "Compact"
vbwProfiler.vbwExecuteLine 7341
           AddPCLproperty "Hardened?", .Hardened, wdBool, "Hardened"
vbwProfiler.vbwExecuteLine 7342
           AddPCLproperty "High Capacity?", .HighCapacity, wdBool, "HighCapacity"
vbwProfiler.vbwExecuteLine 7343
           AddPCLproperty "Dedicated?", .Dedicated, wdBool, "Dedicated"
vbwProfiler.vbwExecuteLine 7344
           AddPCLproperty "Robot Brain?", .RobotBrain, wdBool, "RobotBrain"
vbwProfiler.vbwExecuteLine 7345
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7346
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7347
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7348
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7349
           AddPCLproperty "Complexity", .complexity, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7350
           If .IQ > 0 Then
vbwProfiler.vbwExecuteLine 7351
                AddPCLproperty "IQ", .IQ, wdText, "Disabled"
           End If
vbwProfiler.vbwExecuteLine 7352 'B
vbwProfiler.vbwExecuteLine 7353
           If .DX > 0 Then
vbwProfiler.vbwExecuteLine 7354
                AddPCLproperty "DX", .DX, wdText, "Disabled"
           End If
vbwProfiler.vbwExecuteLine 7355 'B
vbwProfiler.vbwExecuteLine 7356
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7357
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7358
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7359
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7360
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

'vbwLine 7361:    Case Terminal
    Case IIf(vbwProfiler.vbwExecuteLine(7361), VBWPROFILER_EMPTY, _
        Terminal)
vbwProfiler.vbwExecuteLine 7362
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7363
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7364
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7365
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7366
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7367
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7368
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7369
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7370
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7371
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7372
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7373
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7374:    Case DatabaseSoftware
    Case IIf(vbwProfiler.vbwExecuteLine(7374), VBWPROFILER_EMPTY, _
        DatabaseSoftware)
vbwProfiler.vbwExecuteLine 7375
        AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7376
        AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7377
        AddPCLproperty "Size (gigs)", .gigabytes, wdDouble, "Gigabytes"
vbwProfiler.vbwExecuteLine 7378
        AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7379
        AddPCLproperty "Complexity", .complexity, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7380
        AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"

'vbwLine 7381:    Case CartographySoftware, ComputerNavigationSoftware, DatalinkSoftware, TransmissionProfilingSoftware, HoloventureProgram, PersonalitySimulationSoftwareFull, PersonalitySimulationLimited, RoutineVehicleOperationSoftwarePilot, RoutineVehicleOperationSoftwareOther
    Case IIf(vbwProfiler.vbwExecuteLine(7381), VBWPROFILER_EMPTY, _
        CartographySoftware), ComputerNavigationSoftware, DatalinkSoftware, TransmissionProfilingSoftware, HoloventureProgram, PersonalitySimulationSoftwareFull, PersonalitySimulationLimited, RoutineVehicleOperationSoftwarePilot, RoutineVehicleOperationSoftwareOther
vbwProfiler.vbwExecuteLine 7382
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7383
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7384
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7385
           AddPCLproperty "Complexity", .complexity, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7386
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"


'vbwLine 7387:    Case FireDirectionSoftware, TargetingSoftware, DamageControlSoftware, GunnerSoftware, RobotSkillProgramsPhysical, RobotSkillProgramsMental
    Case IIf(vbwProfiler.vbwExecuteLine(7387), VBWPROFILER_EMPTY, _
        FireDirectionSoftware), TargetingSoftware, DamageControlSoftware, GunnerSoftware, RobotSkillProgramsPhysical, RobotSkillProgramsMental
vbwProfiler.vbwExecuteLine 7388
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7389
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7390
           AddPCLproperty "Bonus Skill", .BonusSkillPoints, wdNumber, "BonusSkillPoints"
vbwProfiler.vbwExecuteLine 7391
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7392
           AddPCLproperty "Total Skill Points", .SkillPoints, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7393
           AddPCLproperty "Complexity", .complexity, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7394
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"


'vbwLine 7395:    Case SurgicalInterface, InterfaceWeb, AutoInterfaceWeb, SocketInterface, NeuralInductionField
    Case IIf(vbwProfiler.vbwExecuteLine(7395), VBWPROFILER_EMPTY, _
        SurgicalInterface), InterfaceWeb, AutoInterfaceWeb, SocketInterface, NeuralInductionField
vbwProfiler.vbwExecuteLine 7396
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7397
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7398
           AddPCLproperty "# of Users", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7399
            If component = SocketInterface Then
vbwProfiler.vbwExecuteLine 7400
               AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7401
               AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7402
               AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7403
               AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7404
               AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
            Else
vbwProfiler.vbwExecuteLine 7405 'B
vbwProfiler.vbwExecuteLine 7406
               AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7407
               AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7408
               AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7409
               AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7410
               AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7411
               AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7412
               AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7413
               AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
            End If
vbwProfiler.vbwExecuteLine 7414 'B

'vbwLine 7415:     Case DeflectorField
     Case IIf(vbwProfiler.vbwExecuteLine(7415), VBWPROFILER_EMPTY, _
        DeflectorField)
vbwProfiler.vbwExecuteLine 7416
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7417
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7418
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7419
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7420
           AddPCLproperty "PD Bonus", "+" + Format(.PDBonus), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7421
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7422
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7423
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           'AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           'AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           'AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

'vbwLine 7424:    Case ForceScreen, VariableForceScreen
    Case IIf(vbwProfiler.vbwExecuteLine(7424), VBWPROFILER_EMPTY, _
        ForceScreen), VariableForceScreen
vbwProfiler.vbwExecuteLine 7425
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7426
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7427
           AddPCLproperty "Screen DR", .ForceDR, wdNumber, "ForceDR"
vbwProfiler.vbwExecuteLine 7428
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7429
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7430
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7431
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7432
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
           'AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
           'AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
           'AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

End Select
vbwProfiler.vbwExecuteLine 7433 'B
vbwProfiler.vbwExecuteLine 7434
End With
vbwProfiler.vbwProcOut 277
vbwProfiler.vbwExecuteLine 7435
End Sub

Private Sub ShowPropsForMiscellanous(ByVal component As Integer, ByVal Key As String)
'///////////////////////////////////////////
'Miscellanous equipment
'///////////////////////////////////////////
vbwProfiler.vbwProcIn 278
vbwProfiler.vbwExecuteLine 7436
With m_oCurrentVeh.Components(Key)

vbwProfiler.vbwExecuteLine 7437
Select Case component
' Fill the window with properties for the correct Collection item
'vbwLine 7438:    Case ArmMotor
    Case IIf(vbwProfiler.vbwExecuteLine(7438), VBWPROFILER_EMPTY, _
        ArmMotor)
vbwProfiler.vbwExecuteLine 7439
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7440
           AddPCLproperty "ST", .ST, wdNumber, "ST"
vbwProfiler.vbwExecuteLine 7441
           AddPCLproperty "Bad Grip?", .BadGrip, wdBool, "BadGrip"
vbwProfiler.vbwExecuteLine 7442
           AddPCLproperty "Cheap?", .Cheap, wdBool, "Cheap"
vbwProfiler.vbwExecuteLine 7443
           AddPCLproperty "Extendable?", .Extendable, wdBool, "Extendable"
vbwProfiler.vbwExecuteLine 7444
           AddPCLproperty "Poor Coordination?", .PoorCoordination, wdBool, "PoorCoordination"
vbwProfiler.vbwExecuteLine 7445
           AddPCLproperty "Striker?", .Striker, wdBool, "Striker"
vbwProfiler.vbwExecuteLine 7446
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7447
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7448
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7449
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7450
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7451
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7452
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7453
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7454
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7455:    Case FireExtinguisherSystem, FullFireSuppressionSystem, CompactFireSuppressionSystem
    Case IIf(vbwProfiler.vbwExecuteLine(7455), VBWPROFILER_EMPTY, _
        FireExtinguisherSystem), FullFireSuppressionSystem, CompactFireSuppressionSystem
vbwProfiler.vbwExecuteLine 7456
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7457
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7458
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7459
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7460
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7461
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7462
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7463
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7464
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7465
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7466
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

'vbwLine 7467:    Case BilgePump
    Case IIf(vbwProfiler.vbwExecuteLine(7467), VBWPROFILER_EMPTY, _
        BilgePump)
vbwProfiler.vbwExecuteLine 7468
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7469
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7470
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7471
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7472
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7473
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7474
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7475
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7476
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7477
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7478
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7479
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7480:    Case CompleteWorkshop, MechanicWorkshop, EngineeringWorkshop, ElectronicsWorkshop, ArmouryWorkshop, CompleteMiniWorkshop, ScienceLab, MiniMechanicWorkshop, MiniElectronicsWorkshop, MiniEngineeringWorkshop, MiniArmouryWorkshop
    Case IIf(vbwProfiler.vbwExecuteLine(7480), VBWPROFILER_EMPTY, _
        CompleteWorkshop), MechanicWorkshop, EngineeringWorkshop, ElectronicsWorkshop, ArmouryWorkshop, CompleteMiniWorkshop, ScienceLab, MiniMechanicWorkshop, MiniElectronicsWorkshop, MiniEngineeringWorkshop, MiniArmouryWorkshop
vbwProfiler.vbwExecuteLine 7481
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7482
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7483
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7484
            If component = ScienceLab Then
               'AddPCLproperty "Skill", .Skill, wdList, "Skill", "astronomy", "biochemistry", "biology", "botany", "chemistry", "computer programming", "criminology", "ecology", "economics", "electronics", "engineering", "forensics", "genetics", "geology", "history", "linguistics", "literature", "mathematics", "metallurgy", "meteorology", "nuclear physics", "occultism", "physics", "physiology", "prospecting", "psychology", "research", "theology", "zoology"
vbwProfiler.vbwExecuteLine 7485
               AddPCLproperty "Skill", .Skill, wdText, "Skill"
            End If
vbwProfiler.vbwExecuteLine 7486 'B
vbwProfiler.vbwExecuteLine 7487
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7488
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7489
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7490
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7491
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7492
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7493
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7494
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7495
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7496:    Case ExtendableLadder, Crane, Winch, PowerShovel, WreckingCrane, ForkLift, VehicularBridge, LaunchCatapult, SkyHook, Bore, SuperBore, EnergyDrill, TractorBeam, PressorBeam, CombinationBeam, CraneWithElectroMagnet
    Case IIf(vbwProfiler.vbwExecuteLine(7496), VBWPROFILER_EMPTY, _
        ExtendableLadder), Crane, Winch, PowerShovel, WreckingCrane, ForkLift, VehicularBridge, LaunchCatapult, SkyHook, Bore, SuperBore, EnergyDrill, TractorBeam, PressorBeam, CombinationBeam, CraneWithElectroMagnet
vbwProfiler.vbwExecuteLine 7497
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7498
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7499
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7500
            Select Case component
'vbwLine 7501:            Case ExtendableLadder, Crane, CraneWithElectroMagnet, WreckingCrane
            Case IIf(vbwProfiler.vbwExecuteLine(7501), VBWPROFILER_EMPTY, _
        ExtendableLadder), Crane, CraneWithElectroMagnet, WreckingCrane
vbwProfiler.vbwExecuteLine 7502
               AddPCLproperty "Crane Height (ft)", .Height, wdNumber, "Height"
'vbwLine 7503:            Case PowerShovel, Winch, ForkLift, TractorBeam, PressorBeam, CombinationBeam
            Case IIf(vbwProfiler.vbwExecuteLine(7503), VBWPROFILER_EMPTY, _
        PowerShovel), Winch, ForkLift, TractorBeam, PressorBeam, CombinationBeam
vbwProfiler.vbwExecuteLine 7504
               AddPCLproperty "ST", .ST, wdNumber, "ST"
'vbwLine 7505:            Case VehicularBridge
            Case IIf(vbwProfiler.vbwExecuteLine(7505), VBWPROFILER_EMPTY, _
        VehicularBridge)
vbwProfiler.vbwExecuteLine 7506
               AddPCLproperty "Length (yds)", .Length, wdDouble, "Length"
vbwProfiler.vbwExecuteLine 7507
               AddPCLproperty "Max Supported Weight", .DesiredWeight, wdDouble, "DesiredWeight"
'vbwLine 7508:            Case Bore, SuperBore
            Case IIf(vbwProfiler.vbwExecuteLine(7508), VBWPROFILER_EMPTY, _
        Bore), SuperBore
vbwProfiler.vbwExecuteLine 7509
               AddPCLproperty "Tunneling Ability Per Hour (cf)", .TunnelingAbility, wdDouble, "TunnelingAbility"
'vbwLine 7510:            Case SkyHook, LaunchCatapult
            Case IIf(vbwProfiler.vbwExecuteLine(7510), VBWPROFILER_EMPTY, _
        SkyHook), LaunchCatapult
            End Select
vbwProfiler.vbwExecuteLine 7511 'B
vbwProfiler.vbwExecuteLine 7512
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7513
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7514
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7515
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7516
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7517
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7518
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7519
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7520
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7521:    Case OperatingRoom, StretcherPallet, EmergencySupportUnit, EmergencylightsandSiren, CryonicCapsule, Automed, DiagnosisTable
    Case IIf(vbwProfiler.vbwExecuteLine(7521), VBWPROFILER_EMPTY, _
        OperatingRoom), StretcherPallet, EmergencySupportUnit, EmergencylightsandSiren, CryonicCapsule, Automed, DiagnosisTable
vbwProfiler.vbwExecuteLine 7522
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7523
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7524
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7525
            If component = OperatingRoom Then
vbwProfiler.vbwExecuteLine 7526
               AddPCLproperty "# of Operating Tables", .OperatingTables, wdNumber, "OperatingTables"
            End If
vbwProfiler.vbwExecuteLine 7527 'B
vbwProfiler.vbwExecuteLine 7528
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7529
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7530
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7531
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7532
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7533
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7534
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7535
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7536
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7537:    Case Stage, Hall, BarRoom, ConferenceRoom, MovieScreenandProjector, MovieScreenandProjectorSmall, HoloventureZone
    Case IIf(vbwProfiler.vbwExecuteLine(7537), VBWPROFILER_EMPTY, _
        Stage), Hall, BarRoom, ConferenceRoom, MovieScreenandProjector, MovieScreenandProjectorSmall, HoloventureZone
vbwProfiler.vbwExecuteLine 7538
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7539
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7540
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7541
            Select Case component
'vbwLine 7542:                Case Stage, Hall, BarRoom, ConferenceRoom, HoloventureZone
                Case IIf(vbwProfiler.vbwExecuteLine(7542), VBWPROFILER_EMPTY, _
        Stage), Hall, BarRoom, ConferenceRoom, HoloventureZone
vbwProfiler.vbwExecuteLine 7543
                   AddPCLproperty "Floor Area", .FloorArea, wdDouble, "FloorArea"
                Case Else
vbwProfiler.vbwExecuteLine 7544 'B
vbwProfiler.vbwExecuteLine 7545
                    AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
            End Select
vbwProfiler.vbwExecuteLine 7546 'B
vbwProfiler.vbwExecuteLine 7547
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7548
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7549
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7550
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7551
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7552
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7553
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7554
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

        'Note: door and hatch have been remove below.  User should enter these in as "Details" in the options dialog
'vbwLine 7555:    Case CargoRamp, Airlock, MembraneAirlock, Forcelock, PassageTube, ArmoredPassageTube
    Case IIf(vbwProfiler.vbwExecuteLine(7555), VBWPROFILER_EMPTY, _
        CargoRamp), Airlock, MembraneAirlock, Forcelock, PassageTube, ArmoredPassageTube
vbwProfiler.vbwExecuteLine 7556
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7557
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7558
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7559
            Select Case component
'vbwLine 7560:                Case Airlock, MembraneAirlock
                Case IIf(vbwProfiler.vbwExecuteLine(7560), VBWPROFILER_EMPTY, _
        Airlock), MembraneAirlock
vbwProfiler.vbwExecuteLine 7561
                   AddPCLproperty "# People Supported", .Rating, wdNumber, "Rating"
                Case Else
vbwProfiler.vbwExecuteLine 7562 'B
            End Select
vbwProfiler.vbwExecuteLine 7563 'B
vbwProfiler.vbwExecuteLine 7564
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7565
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7566
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7567
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7568
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7569
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7570
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7571
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7572
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7573:    Case TeleportProjector
    Case IIf(vbwProfiler.vbwExecuteLine(7573), VBWPROFILER_EMPTY, _
        TeleportProjector)
vbwProfiler.vbwExecuteLine 7574
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7575
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7576
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7577
           AddPCLproperty "# Hexes", .HexCapacity, wdNumber, "HexCapacity"
vbwProfiler.vbwExecuteLine 7578
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7579
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7580
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7581
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7582
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7583
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7584
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7585
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7586
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7587:    Case BrigsandRestraints, BurglarAlarm, HighSecurityAlarm, MutableLicensePlate, OilSprayer, PaintSprayer, SmokeScreen, SpikeDropper
    Case IIf(vbwProfiler.vbwExecuteLine(7587), VBWPROFILER_EMPTY, _
        BrigsandRestraints), BurglarAlarm, HighSecurityAlarm, MutableLicensePlate, OilSprayer, PaintSprayer, SmokeScreen, SpikeDropper
vbwProfiler.vbwExecuteLine 7588
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7589
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7590
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7591
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7592
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7593
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7594
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7595
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7596
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7597
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7598
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7599
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7600:    Case VehicleBay, HangerBay, DryDock, SpaceDock, ExternalCradle
    Case IIf(vbwProfiler.vbwExecuteLine(7600), VBWPROFILER_EMPTY, _
        VehicleBay), HangerBay, DryDock, SpaceDock, ExternalCradle
vbwProfiler.vbwExecuteLine 7601
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7602
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7603
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7604
            If component = ExternalCradle Then
vbwProfiler.vbwExecuteLine 7605
               AddPCLproperty "Total Craft Weight", .CraftWeight, wdDouble, "CraftWeight"
            Else
vbwProfiler.vbwExecuteLine 7606 'B
vbwProfiler.vbwExecuteLine 7607
               AddPCLproperty "Total Craft Weight", .CraftWeight, wdDouble, "CraftWeight"
vbwProfiler.vbwExecuteLine 7608
               AddPCLproperty "Cubic Feet of Craft", .CubicFeetCraft, wdDouble, "CubicFeetCraft"
            End If
vbwProfiler.vbwExecuteLine 7609 'B
vbwProfiler.vbwExecuteLine 7610
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7611
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7612
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7613
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7614
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7615
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7616
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7617
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7618:    Case ArrestorHook, VehicularParachute
    Case IIf(vbwProfiler.vbwExecuteLine(7618), VBWPROFILER_EMPTY, _
        ArrestorHook), VehicularParachute
vbwProfiler.vbwExecuteLine 7619
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7620
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7621
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7622
            If component = VehicularParachute Then
vbwProfiler.vbwExecuteLine 7623
               AddPCLproperty "Rated Weight", .RatedWeight, wdDouble, "RatedWeight"
            End If
vbwProfiler.vbwExecuteLine 7624 'B
vbwProfiler.vbwExecuteLine 7625
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7626
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7627
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7628
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7629
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7630
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7631
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7632
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7633:    Case RefuellingProbe, RefuellingDrogue, FuelElectrolysisSystem, HydrogenFuelScoop, AtmosphereProcessor
    Case IIf(vbwProfiler.vbwExecuteLine(7633), VBWPROFILER_EMPTY, _
        RefuellingProbe), RefuellingDrogue, FuelElectrolysisSystem, HydrogenFuelScoop, AtmosphereProcessor
vbwProfiler.vbwExecuteLine 7634
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7635
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7636
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7637
            If (component = FuelElectrolysisSystem) Or (component = AtmosphereProcessor) Then
vbwProfiler.vbwExecuteLine 7638
               AddPCLproperty "Processing Capacity (gallons)", .capacity, wdDouble, "Capacity"
            End If
vbwProfiler.vbwExecuteLine 7639 'B
vbwProfiler.vbwExecuteLine 7640
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7641
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7642
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7643
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7644
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7645
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7646
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7647
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7648
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7649:    Case NuclearDamper
    Case IIf(vbwProfiler.vbwExecuteLine(7649), VBWPROFILER_EMPTY, _
        NuclearDamper)
vbwProfiler.vbwExecuteLine 7650
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7651
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7652
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7653
           AddPCLproperty "Field Radius (mi)", .Radius, wdDouble, "Radius"
vbwProfiler.vbwExecuteLine 7654
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7655
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7656
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7657
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7658
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7659
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7660
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7661
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7662
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7663:    Case SmallRealityStabilizer, MediumRealityStabilizer, HeavyRealityStabilizer
    Case IIf(vbwProfiler.vbwExecuteLine(7663), VBWPROFILER_EMPTY, _
        SmallRealityStabilizer), MediumRealityStabilizer, HeavyRealityStabilizer
vbwProfiler.vbwExecuteLine 7664
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7665
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7666
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7667
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7668
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7669
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7670
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7671
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7672
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7673
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7674
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7675
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7676:    Case ModularSocket
    Case IIf(vbwProfiler.vbwExecuteLine(7676), VBWPROFILER_EMPTY, _
        ModularSocket)
vbwProfiler.vbwExecuteLine 7677
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7678
           AddPCLproperty "Rated Volume", .RatedVolume, wdDouble, "RatedVolume"
vbwProfiler.vbwExecuteLine 7679
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7680
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7681
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"


'vbwLine 7682:    Case Module
    Case IIf(vbwProfiler.vbwExecuteLine(7682), VBWPROFILER_EMPTY, _
        Module)
vbwProfiler.vbwExecuteLine 7683
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7684
           AddPCLproperty "Waste Weight", .WasteWeight, wdDouble, "WasteWeight"
vbwProfiler.vbwExecuteLine 7685
           AddPCLproperty "Waste Volume", .WasteVolume, wdDouble, "WasteVolume"
vbwProfiler.vbwExecuteLine 7686
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7687
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7688
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7689
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"



End Select
vbwProfiler.vbwExecuteLine 7690 'B
vbwProfiler.vbwExecuteLine 7691
End With
vbwProfiler.vbwProcOut 278
vbwProfiler.vbwExecuteLine 7692
End Sub
        
Private Sub ShowPropsForPowerandFuel(ByVal component As Integer, ByVal Key As String)
'///////////////////////////////////////////
'Power and Fuel
'//////////////////////////////////////////
vbwProfiler.vbwProcIn 279


Dim listarray() As String
vbwProfiler.vbwExecuteLine 7693
ReDim listarray(1)

vbwProfiler.vbwExecuteLine 7694
With m_oCurrentVeh.Components(Key)

' Fill the window with properties for the correct Collection item

vbwProfiler.vbwExecuteLine 7695
Select Case component

'vbwLine 7696: Case MuscleEngine
 Case IIf(vbwProfiler.vbwExecuteLine(7696), VBWPROFILER_EMPTY, _
        MuscleEngine)
vbwProfiler.vbwExecuteLine 7697
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7698
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7699
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7700
           AddPCLproperty "Maximum Output", .MaxOutput, wdDouble, "MaxOutPut"
vbwProfiler.vbwExecuteLine 7701
           AddPCLproperty "Combined Operator ST", .CombinedST, wdNumber, "CombinedST"
vbwProfiler.vbwExecuteLine 7702
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7703
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7704
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7705
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7706
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7707
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7708
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7709
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7710
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7711
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7712
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7713:Case EarlySteamEngine, ForcedDraftSteamEngine, TripleExpansionSteamEngine, SteamTurbine
Case IIf(vbwProfiler.vbwExecuteLine(7713), VBWPROFILER_EMPTY, _
        EarlySteamEngine), ForcedDraftSteamEngine, TripleExpansionSteamEngine, SteamTurbine
vbwProfiler.vbwExecuteLine 7714
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7715
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7716
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7717
           AddPCLproperty "Output", .DesiredOutput, wdDouble, "DesiredOutput"
vbwProfiler.vbwExecuteLine 7718
           If component = SteamTurbine Then
vbwProfiler.vbwExecuteLine 7719
                AddPCLproperty "Fuel Type", .Fueltype, wdList, "FuelType", "coal", "diesel fuel"
            Else
vbwProfiler.vbwExecuteLine 7720 'B
vbwProfiler.vbwExecuteLine 7721
                AddPCLproperty "Fuel Type", .Fueltype, wdList, "FuelType", "coal", "wood"
            End If
vbwProfiler.vbwExecuteLine 7722 'B

vbwProfiler.vbwExecuteLine 7723
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7724
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7725
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7726
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7727
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7728
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7729
           AddPCLproperty "Fuel Consumption", .FuelConsumption, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7730
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7731
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7732
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7733
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7734
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7735: Case GasolineEngine, HPGasolineEngine, TurboGasolineEngine, SuperGasolineEngine, TurboHPGasolineEngine, SuperHPGasolineEngine, StandardDieselEngine, TurboStandardDieselEngine, MarineDieselEngine, HPDieselEngine, TurboHPDieselEngine, CeramicEngine, TurboCeramicEngine, SuperCeramicEngine, HPCeramicEngine, TurboHPCeramicEngine, SuperHPCeramicEngine, HydrogenCombustionEngine
 Case IIf(vbwProfiler.vbwExecuteLine(7735), VBWPROFILER_EMPTY, _
        GasolineEngine), HPGasolineEngine, TurboGasolineEngine, SuperGasolineEngine, TurboHPGasolineEngine, SuperHPGasolineEngine, StandardDieselEngine, TurboStandardDieselEngine, MarineDieselEngine, HPDieselEngine, TurboHPDieselEngine, CeramicEngine, TurboCeramicEngine, SuperCeramicEngine, HPCeramicEngine, TurboHPCeramicEngine, SuperHPCeramicEngine, HydrogenCombustionEngine
vbwProfiler.vbwExecuteLine 7736
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7737
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7738
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7739
           AddPCLproperty "Output", .UnModifiedOutput, wdDouble, "UnModifiedOutput"
           'hydrogen combustion
vbwProfiler.vbwExecuteLine 7740
            If component = HydrogenCombustionEngine Then
vbwProfiler.vbwExecuteLine 7741
                AddPCLproperty "Fuel Type", .Fueltype, wdList, "FuelType", "hydrogen"
            'aviation fuels
'vbwLine 7742:            ElseIf (component = HPCeramicEngine) Or (component = TurboHPCeramicEngine) Or (component = SuperHPCeramicEngine) Or (component = HPGasolineEngine) Or (component = TurboHPGasolineEngine) Or (component = SuperHPGasolineEngine) Then
            ElseIf vbwProfiler.vbwExecuteLine(7742) Or (component = HPCeramicEngine) Or (component = TurboHPCeramicEngine) Or _
                (component = SuperHPCeramicEngine) Or (component = HPGasolineEngine) Or _
                (component = TurboHPGasolineEngine) Or (component = SuperHPGasolineEngine) Then
vbwProfiler.vbwExecuteLine 7743
                AddPCLproperty "Fuel Type", .Fueltype, wdList, "FuelType", "aviation gas"
            'multifuels
'vbwLine 7744:            ElseIf (component = CeramicEngine) Or (component = TurboCeramicEngine) Or (component = SuperCeramicEngine) Then
            ElseIf vbwProfiler.vbwExecuteLine(7744) Or (component = CeramicEngine) Or (component = TurboCeramicEngine) Or (component = SuperCeramicEngine) Then
vbwProfiler.vbwExecuteLine 7745
                AddPCLproperty "Fuel Type", .Fueltype, wdList, "FuelType", "gasoline", "diesel fuel", "aviation gas", "ethanol", "methanol"
            'diesels with alcohol / propane potential
'vbwLine 7746:            ElseIf (component = TurboStandardDieselEngine) Or (component = TurboHPDieselEngine) Or (component = MarineDieselEngine) Or (component = StandardDieselEngine) Or (component = HPDieselEngine) Then
            ElseIf vbwProfiler.vbwExecuteLine(7746) Or (component = TurboStandardDieselEngine) Or (component = TurboHPDieselEngine) Or _
                (component = MarineDieselEngine) Or (component = StandardDieselEngine) Or _
                (component = HPDieselEngine) Then
vbwProfiler.vbwExecuteLine 7747
                AddPCLproperty "Fuel Type", .Fueltype, wdList, "FuelType", "diesel fuel", "propane", "ethanol", "methanol"
            'gasolines with alcohol / propane potential
'vbwLine 7748:            ElseIf (component = GasolineEngine) Or (component = TurboGasolineEngine) Or (component = SuperGasolineEngine) Then
            ElseIf vbwProfiler.vbwExecuteLine(7748) Or (component = GasolineEngine) Or (component = TurboGasolineEngine) Or _
            (component = SuperGasolineEngine) Then
vbwProfiler.vbwExecuteLine 7749
                AddPCLproperty "Fuel Type", .Fueltype, wdList, "FuelType", "gasoline", "propane", "ethanol", "methanol"
            End If
vbwProfiler.vbwExecuteLine 7750 'B
vbwProfiler.vbwExecuteLine 7751
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7752
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7753
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7754
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7755
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7756
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7757
           AddPCLproperty "Fuel Consumption", .FuelConsumption, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7758
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7759
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7760
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7761
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7762
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"



'vbwLine 7763:Case FuelCell, HPGasTurbine, StandardMHDTurbine, HPMHDTurbine
Case IIf(vbwProfiler.vbwExecuteLine(7763), VBWPROFILER_EMPTY, _
        FuelCell), HPGasTurbine, StandardMHDTurbine, HPMHDTurbine
vbwProfiler.vbwExecuteLine 7764
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7765
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7766
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7767
           AddPCLproperty "Output", .DesiredOutput, wdDouble, "DesiredOutPut"
vbwProfiler.vbwExecuteLine 7768
           AddPCLproperty "Closed Cycle?", .ClosedCycle, wdBool, "ClosedCycle"
vbwProfiler.vbwExecuteLine 7769
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7770
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7771
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7772
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7773
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7774
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7775
           AddPCLproperty "Fuel Consumption", .FuelConsumption, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7776
           AddPCLproperty "LOX Consumption", .LOXConsumption, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7777
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7778
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7779
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7780
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7781
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

'vbwLine 7782:Case StandardGasTurbine, OptimizedGasTurbine
Case IIf(vbwProfiler.vbwExecuteLine(7782), VBWPROFILER_EMPTY, _
        StandardGasTurbine), OptimizedGasTurbine
vbwProfiler.vbwExecuteLine 7783
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7784
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7785
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7786
           AddPCLproperty "Output", .DesiredOutput, wdDouble, "DesiredOutPut"
vbwProfiler.vbwExecuteLine 7787
           AddPCLproperty "Fuel Type", .Fueltype, wdList, "FuelType", "gasoline", "diesel fuel", "alcohol", "aviation gas", "jet fuel"
vbwProfiler.vbwExecuteLine 7788
           AddPCLproperty "Closed Cycle?", .ClosedCycle, wdBool, "ClosedCycle"
vbwProfiler.vbwExecuteLine 7789
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7790
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7791
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7792
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7793
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7794
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7795
           AddPCLproperty "Fuel Consumption", .FuelConsumption, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7796
           AddPCLproperty "LOX Consumption", .LOXConsumption, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7797
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7798
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7799
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7800
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7801
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


    'only difference between FissionReactor and the others is the Uranium Fuel Rods
'vbwLine 7802:Case FissionReactor
Case IIf(vbwProfiler.vbwExecuteLine(7802), VBWPROFILER_EMPTY, _
        FissionReactor)
vbwProfiler.vbwExecuteLine 7803
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7804
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7805
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7806
           AddPCLproperty "Output", .DesiredOutput, wdDouble, "DesiredOutput"
vbwProfiler.vbwExecuteLine 7807
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7808
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7809
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7810
           AddPCLproperty "Endurance", .Endurance & " yrs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7811
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7812
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7813
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7814
           AddPCLproperty "Fuel Rods Installed", .FuelConsumption, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7815
           AddPCLproperty "Fuel Rod Added Cost", .FuelCost, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7816
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7817
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7818
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7819
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7820
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"



'vbwLine 7821:Case RTGReactor, NPU, FusionReactor, AntimatterReactor, TotalConversionPowerPlant, CosmicPowerPlant
Case IIf(vbwProfiler.vbwExecuteLine(7821), VBWPROFILER_EMPTY, _
        RTGReactor), NPU, FusionReactor, AntimatterReactor, TotalConversionPowerPlant, CosmicPowerPlant
vbwProfiler.vbwExecuteLine 7822
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7823
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7824
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7825
           AddPCLproperty "Output", .DesiredOutput, wdDouble, "DesiredOutPut"
vbwProfiler.vbwExecuteLine 7826
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7827
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7828
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7829
           AddPCLproperty "Endurance", .Endurance & " yrs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7830
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7831
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7832
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7833
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7834
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7835
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7836
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7837
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"



'vbwLine 7838:Case Soulburner, ElementalFurnace, ManaEngine
Case IIf(vbwProfiler.vbwExecuteLine(7838), VBWPROFILER_EMPTY, _
        Soulburner), ElementalFurnace, ManaEngine
vbwProfiler.vbwExecuteLine 7839
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7840
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7841
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7842
           AddPCLproperty "Output", .DesiredOutput, wdDouble, "DesiredOutPut"
vbwProfiler.vbwExecuteLine 7843
           AddPCLproperty "Cost for Magic", .MagicCost, wdDouble, "MagicCost"
vbwProfiler.vbwExecuteLine 7844
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7845
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7846
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7847
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7848
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7849
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7850
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7851
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7852
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7853
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7854
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7855:Case Carnivore, Herbivore, Omnivore, Vampire
Case IIf(vbwProfiler.vbwExecuteLine(7855), VBWPROFILER_EMPTY, _
        Carnivore), Herbivore, Omnivore, Vampire
vbwProfiler.vbwExecuteLine 7856
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7857
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7858
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7859
           AddPCLproperty "Output", .DesiredOutput, wdDouble, "DesiredOutPut"
vbwProfiler.vbwExecuteLine 7860
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7861
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7862
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7863
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7864
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7865
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7866
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7867
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7868
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7869
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7870
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7871:Case ClockWork
Case IIf(vbwProfiler.vbwExecuteLine(7871), VBWPROFILER_EMPTY, _
        ClockWork)
vbwProfiler.vbwExecuteLine 7872
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7873
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7874
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7875
           AddPCLproperty "Stored Capacity (kWs)", .DesiredOutput, wdDouble, "DesiredOutPut"
vbwProfiler.vbwExecuteLine 7876
           AddPCLproperty "Powered Rewind Mechanism?", .PoweredRewinder, wdBool, "PoweredRewinder"
vbwProfiler.vbwExecuteLine 7877
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7878
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7879
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7880
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7881
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7882
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7883
           AddPCLproperty "Rewind Motor ST", .MotorST, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7884
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7885
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7886
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7887
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7888
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7889:Case LeadAcidBattery, AdvancedBattery, Flywheel, RechargeablePowerCell, PowerCell
Case IIf(vbwProfiler.vbwExecuteLine(7889), VBWPROFILER_EMPTY, _
        LeadAcidBattery), AdvancedBattery, Flywheel, RechargeablePowerCell, PowerCell
vbwProfiler.vbwExecuteLine 7890
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7891
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7892
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7893
           AddPCLproperty "Stored Capacity (kWs)", .DesiredOutput, wdDouble, "DesiredOutPut"
vbwProfiler.vbwExecuteLine 7894
           If (component = PowerCell) Or (component = RechargeablePowerCell) Then
vbwProfiler.vbwExecuteLine 7895
                AddPCLproperty "Cell Type", .CellType, wdList, "CellType", "custom", "AA", "A", "B", "C", "D", "E"
           End If
vbwProfiler.vbwExecuteLine 7896 'B
vbwProfiler.vbwExecuteLine 7897
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7898
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7899
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7900
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7901
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7902
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7903
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7904
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7905
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7906
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7907
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7908:Case AntiMatterBay
Case IIf(vbwProfiler.vbwExecuteLine(7908), VBWPROFILER_EMPTY, _
        AntiMatterBay)
vbwProfiler.vbwExecuteLine 7909
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7910
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7911
           AddPCLproperty "Capacity (grams)", .capacity, wdDouble, "Capacity"
vbwProfiler.vbwExecuteLine 7912
           AddPCLproperty "Failsafe Points", .FailSafePoints, wdNumber, "FailSafePoints"
vbwProfiler.vbwExecuteLine 7913
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7914
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7915
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7916
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7917
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7918
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7919
           AddPCLproperty "Fuel Cost", .FuelCost, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7920
           AddPCLproperty "Fuel Weight", .FuelWeight, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7921
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7922
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7923:Case StandardTank, lightTank, UltralightTank, StandardSelfSealingTank, lightSelfSealingTank, UltralightSelfSealingTank
Case IIf(vbwProfiler.vbwExecuteLine(7923), VBWPROFILER_EMPTY, _
        StandardTank), lightTank, UltralightTank, StandardSelfSealingTank, lightSelfSealingTank, UltralightSelfSealingTank
vbwProfiler.vbwExecuteLine 7924
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7925
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7926
           AddPCLproperty "Capacity (gallons)", .capacity, wdDouble, "Capacity"
vbwProfiler.vbwExecuteLine 7927
           AddPCLproperty "Fuel Type", .Fuel, wdList, "Fuel", "ethanol", "methanol", "aviation gas", "cadmium", "diesel", "gasoline", "jet fuel", "rocket fuel", "water", "hydrogen", "metal/LOX", "oxygen (LOX)", "propane/LNG"
vbwProfiler.vbwExecuteLine 7928
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7929
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7930
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7931
           AddPCLproperty "Fire", .Fire, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7932
           AddPCLproperty "Tank Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7933
           AddPCLproperty "Tank Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7934
           AddPCLproperty "Tank Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7935
           AddPCLproperty "Fuel Fire", .FuelFire, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7936
           AddPCLproperty "Fuel Cost", .FuelCost, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7937
           AddPCLproperty "Fuel Weight", .FuelWeight, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7938
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7939
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7940:Case CoalBunker, WoodBunker
Case IIf(vbwProfiler.vbwExecuteLine(7940), VBWPROFILER_EMPTY, _
        CoalBunker), WoodBunker
vbwProfiler.vbwExecuteLine 7941
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7942
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7943
           AddPCLproperty "Capacity (cubic ft.)", .capacity, wdDouble, "Capacity"
vbwProfiler.vbwExecuteLine 7944
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7945
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7946
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7947
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7948
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7949
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7950
           AddPCLproperty "Fuel Cost", .FuelCost, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7951
           AddPCLproperty "Fuel Weight", .FuelWeight, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7952
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7953
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'Case Water, Wood, Coal, Gasoline, Diesel, AviationGas, JetFuel, Propane, LiquifiedNaturalGas, EthanolAlchohol, MethanolAlchohol, LiquidHydrogen, LiquidOxygen, Cadmium, MetalLOX, RocketFuel, AntiMatter

'vbwLine 7954:Case ElectricContactPower
Case IIf(vbwProfiler.vbwExecuteLine(7954), VBWPROFILER_EMPTY, _
        ElectricContactPower)
vbwProfiler.vbwExecuteLine 7955
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7956
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7957
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7958
           AddPCLproperty "Power Drawn", .DesiredOutput, wdDouble, "DesiredOutPut"
vbwProfiler.vbwExecuteLine 7959
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7960
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7961
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7962
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7963
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7964
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7965
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7966
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7967
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7968
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7969
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 7970:Case LaserBeamedPowerReceiver, MaserBeamedPowerReceiver
Case IIf(vbwProfiler.vbwExecuteLine(7970), VBWPROFILER_EMPTY, _
        LaserBeamedPowerReceiver), MaserBeamedPowerReceiver
vbwProfiler.vbwExecuteLine 7971
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7972
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 7973
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 7974
           AddPCLproperty "Max Power", .DesiredOutput, wdDouble, "DesiredOutPut"
vbwProfiler.vbwExecuteLine 7975
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7976
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7977
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7978
           AddPCLproperty "Power Output", Format(.Output, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7979
           AddPCLproperty "Consumed", Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7980
           AddPCLproperty "Remaining", Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7981
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7982
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7983
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7984
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7985
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

'vbwLine 7986:Case NitrousOxideBooster
Case IIf(vbwProfiler.vbwExecuteLine(7986), VBWPROFILER_EMPTY, _
        NitrousOxideBooster)
vbwProfiler.vbwExecuteLine 7987
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7988
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 7989
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 7990
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7991
           AddPCLproperty "Max Boost Length", .MaxBoostLength & " seconds", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7992
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7993
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7994
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7995
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 7996
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

'vbwLine 7997:Case Snorkel
Case IIf(vbwProfiler.vbwExecuteLine(7997), VBWPROFILER_EMPTY, _
        Snorkel)

vbwProfiler.vbwExecuteLine 7998
           listarray = .FillCombustionEngineList
vbwProfiler.vbwExecuteLine 7999
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8000
           AddPCLproperty "Assigned Power Plants", .PowerPlants, wdList, "PowerPlants", listarray
vbwProfiler.vbwExecuteLine 8001
           AddPCLproperty "Ruggedized", .Ruggedized, wdBool, "Ruggedized"
vbwProfiler.vbwExecuteLine 8002
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 8003
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8004
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8005
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8006
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8007
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8008
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

End Select
vbwProfiler.vbwExecuteLine 8009 'B

vbwProfiler.vbwExecuteLine 8010
End With
vbwProfiler.vbwProcOut 279
vbwProfiler.vbwExecuteLine 8011
End Sub
   
Private Sub ShowPropsForArmor(ByVal component As Integer, ByVal ComponentsParent As Integer, ByVal Key As String)
vbwProfiler.vbwProcIn 280
Dim listarray() As String

vbwProfiler.vbwExecuteLine 8012
With m_oCurrentVeh.Components(Key)

' Fill the window with properties for the correct Collection item

    '///////////////////////////////////////////
    'Armor
vbwProfiler.vbwExecuteLine 8013
     Select Case component

'vbwLine 8014:        Case ArmorBasicFacing
        Case IIf(vbwProfiler.vbwExecuteLine(8014), VBWPROFILER_EMPTY, _
        ArmorBasicFacing)
vbwProfiler.vbwExecuteLine 8015
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8016
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 8017
           listarray = .FillMaterial
vbwProfiler.vbwExecuteLine 8018
           AddPCLproperty "Material", .material, wdList, "Material", listarray
vbwProfiler.vbwExecuteLine 8019
           listarray = .FillQuality(.material)
vbwProfiler.vbwExecuteLine 8020
           AddPCLproperty "Quality", .Quality, wdList, "Quality", listarray
vbwProfiler.vbwExecuteLine 8021
           AddPCLproperty "Surface Coating", .coating, wdList, "Coating", "none", "reflective", "retro-reflective"
vbwProfiler.vbwExecuteLine 8022
           AddPCLproperty "Radiation Shielding", .radiation, wdBool, "Radiation"
vbwProfiler.vbwExecuteLine 8023
           AddPCLproperty "Thermal Superconductor", .thermal, wdBool, "Thermal"
vbwProfiler.vbwExecuteLine 8024
           AddPCLproperty "Reactive Armor Plating", .rap, wdBool, "RAP"
vbwProfiler.vbwExecuteLine 8025
           AddPCLproperty "Electrified", .electrified, wdBool, "Electrified"
vbwProfiler.vbwExecuteLine 8026
           AddPCLproperty "DR (Right)", .dr1, wdNumber, "DR1"
vbwProfiler.vbwExecuteLine 8027
           AddPCLproperty "DR (Left)", .dr2, wdNumber, "DR2"
vbwProfiler.vbwExecuteLine 8028
           AddPCLproperty "DR (Front)", .dr3, wdNumber, "DR3"
vbwProfiler.vbwExecuteLine 8029
           AddPCLproperty "DR (Back)", .dr4, wdNumber, "DR4"
vbwProfiler.vbwExecuteLine 8030
           AddPCLproperty "DR (Top)", .dr5, wdNumber, "DR5"
vbwProfiler.vbwExecuteLine 8031
           If ComponentsParent = Body Then
vbwProfiler.vbwExecuteLine 8032
               AddPCLproperty "DR (Bottom)", .dr6, wdNumber, "DR6"
           End If
vbwProfiler.vbwExecuteLine 8033 'B
vbwProfiler.vbwExecuteLine 8034
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8035
           AddPCLproperty "Average DR", .AverageDR, wdText, "Disabled"  'MAKE SURE IM Calcing this properly IN THE CLASS depending on whether im dealling with 5 sides or 6!
vbwProfiler.vbwExecuteLine 8036
           AddPCLproperty "Effective DR (Right)", .EffectiveDR1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8037
           AddPCLproperty "Effective DR (Left)", .EffectiveDR2, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8038
           AddPCLproperty "Effective DR (Front)", .EffectiveDR3, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8039
           AddPCLproperty "Effective DR (Back)", .EffectiveDR4, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8040
           AddPCLproperty "Effective DR (Top)", .EffectiveDR5, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8041
           If ComponentsParent = Body Then
vbwProfiler.vbwExecuteLine 8042
               AddPCLproperty "Effective DR (Bottom)", .EffectiveDR6, wdText, "Disabled"
           End If
vbwProfiler.vbwExecuteLine 8043 'B
vbwProfiler.vbwExecuteLine 8044
           AddPCLproperty "PD (Right)", .PD1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8045
           AddPCLproperty "PD (Left)", .PD2, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8046
           AddPCLproperty "PD (Front)", .PD3, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8047
           AddPCLproperty "PD (Back)", .PD4, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8048
           AddPCLproperty "PD (Top)", .PD5, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8049
           If ComponentsParent = Body Then
vbwProfiler.vbwExecuteLine 8050
               AddPCLproperty "PD (Bottom)", .PD6, wdText, "Disabled"
           End If
vbwProfiler.vbwExecuteLine 8051 'B
vbwProfiler.vbwExecuteLine 8052
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8053
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"



'vbwLine 8054:        Case ArmorComplexFacing
        Case IIf(vbwProfiler.vbwExecuteLine(8054), VBWPROFILER_EMPTY, _
        ArmorComplexFacing)
vbwProfiler.vbwExecuteLine 8055
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8056
           AddPCLproperty "TL", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 8057
           AddPCLproperty "Surface Coating", .coating, wdList, "Coating", "none", "reflective", "retro-reflective"
vbwProfiler.vbwExecuteLine 8058
           AddPCLproperty "Radiation Shielding", .radiation, wdBool, "Radiation"
vbwProfiler.vbwExecuteLine 8059
           AddPCLproperty "Thermal Superconductor", .thermal, wdBool, "Thermal"
vbwProfiler.vbwExecuteLine 8060
           AddPCLproperty "Reactive Armor Plating", .rap, wdBool, "RAP"
vbwProfiler.vbwExecuteLine 8061
           AddPCLproperty "Electrified", .electrified, wdBool, "Electrified"
vbwProfiler.vbwExecuteLine 8062
           listarray = .FillMaterial 'only needs to be filled once for all sides
vbwProfiler.vbwExecuteLine 8063
           AddPCLproperty "Material (Right)", .material1, wdList, "Material1", listarray
vbwProfiler.vbwExecuteLine 8064
           AddPCLproperty "Material (Left)", .material2, wdList, "Material2", listarray
vbwProfiler.vbwExecuteLine 8065
           AddPCLproperty "Material (Front)", .material3, wdList, "Material3", listarray
vbwProfiler.vbwExecuteLine 8066
           AddPCLproperty "Material (Back)", .material4, wdList, "Material4", listarray
vbwProfiler.vbwExecuteLine 8067
           AddPCLproperty "Material (Top)", .material5, wdList, "Material5", listarray
vbwProfiler.vbwExecuteLine 8068
           If ComponentsParent = Body Then
vbwProfiler.vbwExecuteLine 8069
               AddPCLproperty "Material (Bottom)", .material6, wdList, "Material6", listarray
           End If
vbwProfiler.vbwExecuteLine 8070 'B
vbwProfiler.vbwExecuteLine 8071
           listarray = .FillQuality(.material1)
vbwProfiler.vbwExecuteLine 8072
           AddPCLproperty "Quality (Right)", .Quality1, wdList, "Quality1", listarray
vbwProfiler.vbwExecuteLine 8073
           listarray = .FillQuality(.material2)
vbwProfiler.vbwExecuteLine 8074
           AddPCLproperty "Quality (Left)", .Quality2, wdList, "Quality2", listarray
vbwProfiler.vbwExecuteLine 8075
           listarray = .FillQuality(.material3)
vbwProfiler.vbwExecuteLine 8076
           AddPCLproperty "Quality (Front)", .Quality3, wdList, "Quality3", listarray
vbwProfiler.vbwExecuteLine 8077
           listarray = .FillQuality(.material4)
vbwProfiler.vbwExecuteLine 8078
           AddPCLproperty "Quality (Back)", .Quality4, wdList, "Quality4", listarray
vbwProfiler.vbwExecuteLine 8079
           listarray = .FillQuality(.material5)
vbwProfiler.vbwExecuteLine 8080
           AddPCLproperty "Quality (Top)", .Quality5, wdList, "Quality5", listarray
vbwProfiler.vbwExecuteLine 8081
           If ComponentsParent = Body Then
vbwProfiler.vbwExecuteLine 8082
                listarray = .FillQuality(.material6)
vbwProfiler.vbwExecuteLine 8083
               AddPCLproperty "Quality (Bottom)", .Quality6, wdList, "Quality6", listarray
           End If
vbwProfiler.vbwExecuteLine 8084 'B
vbwProfiler.vbwExecuteLine 8085
           AddPCLproperty "DR (Right)", .dr1, wdNumber, "DR1", listarray
vbwProfiler.vbwExecuteLine 8086
           AddPCLproperty "DR (Left)", .dr2, wdNumber, "DR2", listarray
vbwProfiler.vbwExecuteLine 8087
           AddPCLproperty "DR (Front)", .dr3, wdNumber, "DR3", listarray
vbwProfiler.vbwExecuteLine 8088
           AddPCLproperty "DR (Back)", .dr4, wdNumber, "DR4", listarray
vbwProfiler.vbwExecuteLine 8089
           AddPCLproperty "DR (Top)", .dr5, wdNumber, "DR5", listarray
vbwProfiler.vbwExecuteLine 8090
           If ComponentsParent = Body Then
vbwProfiler.vbwExecuteLine 8091
                AddPCLproperty "DR (Bottom)", .dr6, wdNumber, "DR6", listarray
           End If
vbwProfiler.vbwExecuteLine 8092 'B
vbwProfiler.vbwExecuteLine 8093
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8094
           AddPCLproperty "Average DR", .AverageDR, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8095
           AddPCLproperty "Effective DR (Right)", .EffectiveDR1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8096
           AddPCLproperty "Effective DR (Left)", .EffectiveDR2, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8097
           AddPCLproperty "Effective DR (Front)", .EffectiveDR3, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8098
           AddPCLproperty "Effective DR (Back)", .EffectiveDR4, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8099
           AddPCLproperty "Effective DR (Top)", .EffectiveDR5, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8100
           If ComponentsParent = Body Then
vbwProfiler.vbwExecuteLine 8101
               AddPCLproperty "Effective DR (Bottom)", .EffectiveDR6, wdText, "Disabled"
           End If
vbwProfiler.vbwExecuteLine 8102 'B
vbwProfiler.vbwExecuteLine 8103
           AddPCLproperty "PD (Right)", .PD1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8104
           AddPCLproperty "PD (Left)", .PD2, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8105
           AddPCLproperty "PD (Front)", .PD3, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8106
           AddPCLproperty "PD (Back)", .PD4, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8107
           AddPCLproperty "PD (Top)", .PD5, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8108
           If ComponentsParent = Body Then
vbwProfiler.vbwExecuteLine 8109
               AddPCLproperty "PD (Bottom)", .PD6, wdText, "Disabled"
           End If
vbwProfiler.vbwExecuteLine 8110 'B
vbwProfiler.vbwExecuteLine 8111
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8112
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"


'vbwLine 8113:        Case ArmorComponent
        Case IIf(vbwProfiler.vbwExecuteLine(8113), VBWPROFILER_EMPTY, _
        ArmorComponent)
vbwProfiler.vbwExecuteLine 8114
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8115
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 8116
           listarray = .FillMaterial
vbwProfiler.vbwExecuteLine 8117
           AddPCLproperty "Material", .material, wdList, "Material", listarray
vbwProfiler.vbwExecuteLine 8118
           listarray = .FillQuality(.material)
vbwProfiler.vbwExecuteLine 8119
           AddPCLproperty "Quality", .Quality, wdList, "Quality", listarray
vbwProfiler.vbwExecuteLine 8120
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 8121
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8122
            AddPCLproperty "PD", .PD, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8123
            AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8124
            AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"


'vbwLine 8125:        Case ArmorLocation
        Case IIf(vbwProfiler.vbwExecuteLine(8125), VBWPROFILER_EMPTY, _
        ArmorLocation)
vbwProfiler.vbwExecuteLine 8126
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8127
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 8128
           listarray = .FillMaterial
vbwProfiler.vbwExecuteLine 8129
           AddPCLproperty "Material", .material, wdList, "Material", listarray
vbwProfiler.vbwExecuteLine 8130
           listarray = .FillQuality(.material)
vbwProfiler.vbwExecuteLine 8131
           AddPCLproperty "Quality", .Quality, wdList, "Quality", listarray
vbwProfiler.vbwExecuteLine 8132
           AddPCLproperty "Surface Coating", .coating, wdList, "Coating", "none", "reflective", "retro-reflective"
vbwProfiler.vbwExecuteLine 8133
           AddPCLproperty "Radiation Shielding", .radiation, wdBool, "Radiation"
vbwProfiler.vbwExecuteLine 8134
           AddPCLproperty "Thermal Superconductor", .thermal, wdBool, "Thermal"
vbwProfiler.vbwExecuteLine 8135
           AddPCLproperty "Reactive Armor Plating", .rap, wdBool, "RAP"
vbwProfiler.vbwExecuteLine 8136
           AddPCLproperty "Electrified", .electrified, wdBool, "Electrified"
vbwProfiler.vbwExecuteLine 8137
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 8138
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8139
            If ComponentsParent = Body Then
vbwProfiler.vbwExecuteLine 8140
                AddPCLproperty "PD (Right)", .PD1, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8141
                AddPCLproperty "PD (Left)", .PD2, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8142
                AddPCLproperty "PD (Front)", .PD3, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8143
                AddPCLproperty "PD (Back)", .PD4, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8144
                AddPCLproperty "PD (Top)", .PD5, wdText, "Disabled"
'vbwLine 8145:            ElseIf ComponentsParent = Turret Or ComponentsParent = Popturret Then
            ElseIf vbwProfiler.vbwExecuteLine(8145) Or ComponentsParent = Turret Or ComponentsParent = Popturret Then
vbwProfiler.vbwExecuteLine 8146
                AddPCLproperty "PD", .PD6, wdText, "Disabled"
            Else
vbwProfiler.vbwExecuteLine 8147 'B
vbwProfiler.vbwExecuteLine 8148
                AddPCLproperty "PD", .PD, wdText, "Disabled"
            End If
vbwProfiler.vbwExecuteLine 8149 'B
vbwProfiler.vbwExecuteLine 8150
            AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8151
            AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"

'vbwLine 8152:        Case ArmorOverall, ArmorWheelGuard, ArmorGunShield
        Case IIf(vbwProfiler.vbwExecuteLine(8152), VBWPROFILER_EMPTY, _
        ArmorOverall), ArmorWheelGuard, ArmorGunShield
vbwProfiler.vbwExecuteLine 8153
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8154
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 8155
           listarray = .FillMaterial
vbwProfiler.vbwExecuteLine 8156
           AddPCLproperty "Material", .material, wdList, "Material", listarray
vbwProfiler.vbwExecuteLine 8157
           listarray = .FillQuality(.material)
vbwProfiler.vbwExecuteLine 8158
           AddPCLproperty "Quality", .Quality, wdList, "Quality", listarray
vbwProfiler.vbwExecuteLine 8159
           AddPCLproperty "Surface Coating", .coating, wdList, "Coating", "none", "reflective", "retro-reflective"
vbwProfiler.vbwExecuteLine 8160
           AddPCLproperty "Radiation Shielding", .radiation, wdBool, "Radiation"
vbwProfiler.vbwExecuteLine 8161
           AddPCLproperty "Thermal Superconductor", .thermal, wdBool, "Thermal"
vbwProfiler.vbwExecuteLine 8162
           AddPCLproperty "Reactive Armor Plating", .rap, wdBool, "RAP"
vbwProfiler.vbwExecuteLine 8163
           AddPCLproperty "Electrified", .electrified, wdBool, "Electrified"
vbwProfiler.vbwExecuteLine 8164
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 8165
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8166
           AddPCLproperty "PD", .PD, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8167
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8168
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"

'vbwLine 8169:        Case ArmorOpenFrame
        Case IIf(vbwProfiler.vbwExecuteLine(8169), VBWPROFILER_EMPTY, _
        ArmorOpenFrame)

vbwProfiler.vbwExecuteLine 8170
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8171
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 8172
           listarray = .FillMaterial
vbwProfiler.vbwExecuteLine 8173
           AddPCLproperty "Material", .material, wdList, "Material", listarray
vbwProfiler.vbwExecuteLine 8174
           listarray = .FillQuality(.material)
vbwProfiler.vbwExecuteLine 8175
           AddPCLproperty "Quality", .Quality, wdList, "Quality", listarray
vbwProfiler.vbwExecuteLine 8176
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 8177
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8178
           AddPCLproperty "PD", .PD, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8179
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8180
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"

    End Select
vbwProfiler.vbwExecuteLine 8181 'B
vbwProfiler.vbwExecuteLine 8182
End With
vbwProfiler.vbwProcOut 280
vbwProfiler.vbwExecuteLine 8183
End Sub

Private Sub ShowPropsForMannedVehicles(ByVal component As Integer, ByVal Key As String)
'//////////////////////////////////////////////
'Manned Vehicle Components
'//////////////////////////////////////////////
vbwProfiler.vbwProcIn 281
vbwProfiler.vbwExecuteLine 8184
 With m_oCurrentVeh.Components(Key)

' Fill the window with properties for the correct Collection item

vbwProfiler.vbwExecuteLine 8185
    Select Case component
'vbwLine 8186:        Case PrimitiveManeuverControl
        Case IIf(vbwProfiler.vbwExecuteLine(8186), VBWPROFILER_EMPTY, _
        PrimitiveManeuverControl)
vbwProfiler.vbwExecuteLine 8187
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8188
           AddPCLproperty "Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 8189
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8190
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8191
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8192
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"


'vbwLine 8193:        Case ElectronicDivingControl, ComputerizedDivingControl, MechanicalManeuverControl, ElectronicManeuverControl, ComputerizedManeuverControl, MechanicalDivingControl
        Case IIf(vbwProfiler.vbwExecuteLine(8193), VBWPROFILER_EMPTY, _
        ElectronicDivingControl), ComputerizedDivingControl, MechanicalManeuverControl, ElectronicManeuverControl, ComputerizedManeuverControl, MechanicalDivingControl
vbwProfiler.vbwExecuteLine 8194
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8195
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 8196
           AddPCLproperty "Duplicate?", .duplicate, wdBool, "Duplicate"
vbwProfiler.vbwExecuteLine 8197
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8198
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8199
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8200
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"


'vbwLine 8201:        Case BattlesuitSystem
        Case IIf(vbwProfiler.vbwExecuteLine(8201), VBWPROFILER_EMPTY, _
        BattlesuitSystem)
vbwProfiler.vbwExecuteLine 8202
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8203
           AddPCLproperty "Tech level", .TL, wdList, "TL", 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 8204
           AddPCLproperty "Pilot Weight", .PilotWeight, wdDouble, "PilotWeight"
vbwProfiler.vbwExecuteLine 8205
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 8206
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8207
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8208
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8209
           AddPCLproperty "Volume1", Format(.Volume1, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8210
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8211
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 8212:        Case FormFittingBattleSuitSystem
        Case IIf(vbwProfiler.vbwExecuteLine(8212), VBWPROFILER_EMPTY, _
        FormFittingBattleSuitSystem)
vbwProfiler.vbwExecuteLine 8213
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8214
           AddPCLproperty "Tech level", .TL, wdList, "TL", 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 8215
           AddPCLproperty "Pilot Weight", .PilotWeight, wdDouble, "PilotWeight"
vbwProfiler.vbwExecuteLine 8216
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 8217
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8218
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8219
           AddPCLproperty "Weight (w/out Pilot)", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8220
           AddPCLproperty "Body Volume", Format(.Volume1, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8221
           AddPCLproperty "Turret Volume", Format(.Volume2, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8222
           AddPCLproperty "Arm Volume (each)", Format(.Volume3, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8223
           AddPCLproperty "Leg Volume (each)", Format(.Volume4, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8224
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8225
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 8226:        Case CrampedSeat, NormalSeat, RoomySeat
        Case IIf(vbwProfiler.vbwExecuteLine(8226), VBWPROFILER_EMPTY, _
        CrampedSeat), NormalSeat, RoomySeat
vbwProfiler.vbwExecuteLine 8227
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8228
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 8229
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 8230
           AddPCLproperty "Exposed?", .Exposed, wdBool, "Exposed"
vbwProfiler.vbwExecuteLine 8231
           AddPCLproperty "G-Seat?", .GSeat, wdBool, "GSeat"
vbwProfiler.vbwExecuteLine 8232
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 8233
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8234
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8235
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8236
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8237
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8238
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

'vbwLine 8239:        Case CycleSeat
        Case IIf(vbwProfiler.vbwExecuteLine(8239), VBWPROFILER_EMPTY, _
        CycleSeat)
vbwProfiler.vbwExecuteLine 8240
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8241
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 8242
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 8243
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 8244
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8245
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8246
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8247
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8248
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8249
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

'vbwLine 8250:        Case CrampedStandingRoom, NormalStandingRoom, RoomyStandingRoom
        Case IIf(vbwProfiler.vbwExecuteLine(8250), VBWPROFILER_EMPTY, _
        CrampedStandingRoom), NormalStandingRoom, RoomyStandingRoom
vbwProfiler.vbwExecuteLine 8251
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8252
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 8253
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 8254
           AddPCLproperty "Exposed?", .Exposed, wdBool, "Exposed"
vbwProfiler.vbwExecuteLine 8255
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 8256
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8257
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8258
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8259
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8260
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8261
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

'vbwLine 8262:        Case Hammock, Bunk, SmallGalley
        Case IIf(vbwProfiler.vbwExecuteLine(8262), VBWPROFILER_EMPTY, _
        Hammock), Bunk, SmallGalley
vbwProfiler.vbwExecuteLine 8263
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8264
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 8265
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 8266
           AddPCLproperty "Added Volume", .AddedVolume, wdDouble, "AddedVolume"
vbwProfiler.vbwExecuteLine 8267
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 8268
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8269
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8270
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8271
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8272
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8273
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

'vbwLine 8274:        Case Cabin, LuxuryCabin, Suite, LuxurySuite
        Case IIf(vbwProfiler.vbwExecuteLine(8274), VBWPROFILER_EMPTY, _
        Cabin), LuxuryCabin, Suite, LuxurySuite
vbwProfiler.vbwExecuteLine 8275
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8276
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 8277
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 8278
           AddPCLproperty "Occupancy", .Occupancy, wdNumber, "Occupancy"
vbwProfiler.vbwExecuteLine 8279
           AddPCLproperty "G-Seats?", .GSeat, wdBool, "Gseat"
vbwProfiler.vbwExecuteLine 8280
           AddPCLproperty "Added Volume", .AddedVolume, wdDouble, "AddedVolume"
vbwProfiler.vbwExecuteLine 8281
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 8282
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8283
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8284
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8285
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8286
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8287
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

'vbwLine 8288:        Case CrampedCrewStation, NormalCrewStation, RoomyCrewStation
        Case IIf(vbwProfiler.vbwExecuteLine(8288), VBWPROFILER_EMPTY, _
        CrampedCrewStation), NormalCrewStation, RoomyCrewStation
vbwProfiler.vbwExecuteLine 8289
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8290
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 8291
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 8292
           AddPCLproperty "Assignment", frmNotes, wdObject, "StationFunction"
vbwProfiler.vbwExecuteLine 8293
           AddPCLproperty "Bridge Access Space?", .BridgeAccessSpace, wdBool, "BridgeAccessSpace"
vbwProfiler.vbwExecuteLine 8294
           AddPCLproperty "Exposed?", .Exposed, wdBool, "Exposed"
vbwProfiler.vbwExecuteLine 8295
           AddPCLproperty "G-Seat?", .GSeat, wdBool, "GSeat"
vbwProfiler.vbwExecuteLine 8296
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 8297
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8298
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8299
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8300
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8301
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8302
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

'vbwLine 8303:        Case CycleCrewStation, HarnessCrewStation
        Case IIf(vbwProfiler.vbwExecuteLine(8303), VBWPROFILER_EMPTY, _
        CycleCrewStation), HarnessCrewStation
vbwProfiler.vbwExecuteLine 8304
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8305
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 8306
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 8307
           AddPCLproperty "Assignment", frmNotes, wdObject, "StationFunction"
vbwProfiler.vbwExecuteLine 8308
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 8309
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8310
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8311
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8312
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8313
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8314
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

'vbwLine 8315:        Case ArtificialGravityUnit
        Case IIf(vbwProfiler.vbwExecuteLine(8315), VBWPROFILER_EMPTY, _
        ArtificialGravityUnit)
vbwProfiler.vbwExecuteLine 8316
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8317
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 8318
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 8319
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8320
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8321
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8322
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8323
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8324
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8325
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 8326:        Case EnvironmentalControl, NBCKit, FullLifeSystem, TotalLifeSystem
        Case IIf(vbwProfiler.vbwExecuteLine(8326), VBWPROFILER_EMPTY, _
        EnvironmentalControl), NBCKit, FullLifeSystem, TotalLifeSystem
vbwProfiler.vbwExecuteLine 8327
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8328
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 8329
           AddPCLproperty "# People", .People, wdNumber, "People"
vbwProfiler.vbwExecuteLine 8330
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 8331
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8332
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8333
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8334
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8335
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8336
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8337
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 8338:        Case LimitedLifeSystem
        Case IIf(vbwProfiler.vbwExecuteLine(8338), VBWPROFILER_EMPTY, _
        LimitedLifeSystem)
vbwProfiler.vbwExecuteLine 8339
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8340
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 8341
           AddPCLproperty "# People", .People, wdNumber, "People"
vbwProfiler.vbwExecuteLine 8342
           AddPCLproperty "# Man Days", .ManDays, wdDouble, "ManDays"
vbwProfiler.vbwExecuteLine 8343
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 8344
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8345
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8346
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8347
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8348
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8349
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8350
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 8351:        Case EjectionSeat, CrewEscapeCapsule, Airbag, CrashWeb, WombTank, GravityWeb
        Case IIf(vbwProfiler.vbwExecuteLine(8351), VBWPROFILER_EMPTY, _
        EjectionSeat), CrewEscapeCapsule, Airbag, CrashWeb, WombTank, GravityWeb
vbwProfiler.vbwExecuteLine 8352
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8353
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 8354
           If component = CrewEscapeCapsule Then
vbwProfiler.vbwExecuteLine 8355
                AddPCLproperty "Max Occupancy", .Occupancy, wdNumber, "Occupancy"
           End If
vbwProfiler.vbwExecuteLine 8356 'B
vbwProfiler.vbwExecuteLine 8357
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 8358
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8359
           If component = GravityWeb Then
vbwProfiler.vbwExecuteLine 8360
                AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
           End If
vbwProfiler.vbwExecuteLine 8361 'B
vbwProfiler.vbwExecuteLine 8362
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8363
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8364
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8365
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8366
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"


'vbwLine 8367:        Case GravCompensator
        Case IIf(vbwProfiler.vbwExecuteLine(8367), VBWPROFILER_EMPTY, _
        GravCompensator)
vbwProfiler.vbwExecuteLine 8368
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8369
           AddPCLproperty "Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
vbwProfiler.vbwExecuteLine 8370
           AddPCLproperty "Quantity", .Quantity, wdNumber, "Quantity"
vbwProfiler.vbwExecuteLine 8371
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 8372
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8373
           AddPCLproperty "Power Consumption", Format(.PowerReqt, "standard") & " kW", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8374
           AddPCLproperty "G reduction", .GReduction, wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8375
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8376
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8377
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8378
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8379
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

'vbwLine 8380:        Case Provisions
        Case IIf(vbwProfiler.vbwExecuteLine(8380), VBWPROFILER_EMPTY, _
        Provisions)
vbwProfiler.vbwExecuteLine 8381
           AddPCLproperty "Settings", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8382
           AddPCLproperty "# days worth", .occupancydays, wdNumber, "occupancydays"
vbwProfiler.vbwExecuteLine 8383
           AddPCLproperty "Settings", .Setting, wdList, "Setting", "auto", "light", "heavy"
vbwProfiler.vbwExecuteLine 8384
           AddPCLproperty "DR", .dr, wdNumber, "DR"
vbwProfiler.vbwExecuteLine 8385
           AddPCLproperty "Statistics", "", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8386
           AddPCLproperty "Cost", "$" & Format(.Cost, "standard"), wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8387
           AddPCLproperty "Weight", Format(.Weight, "standard") & " lbs", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8388
           AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8389
           AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
vbwProfiler.vbwExecuteLine 8390
           AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"

 End Select
vbwProfiler.vbwExecuteLine 8391 'B
vbwProfiler.vbwExecuteLine 8392
End With
vbwProfiler.vbwProcOut 281
vbwProfiler.vbwExecuteLine 8393
End Sub


