Attribute VB_Name = "modGUI_Power"
Option Explicit

                    'Private Sub DisplayFuel()
                    'Dim consumption As Single
                    'Dim capacity As Single
                    'Dim item As ListItem
                    'Dim i As Long
                    '
                    '
                    ''get the total capacity
                    'If lstvLocal.ListItems.Count <> 0 Then
                    '    capacity = 0
                    '    For i = 1 To lstvLocal.ListItems.Count
                    '        capacity = capacity + m_oCurrentVeh.Components(lstvLocal.ListItems(i).Key).capacity
                    '    Next
                    'Else
                    '    capacity = 0
                    'End If
                    '
                    ''round the value for consumed
                    'lstvSystem.ListItems.item(msCurrentSystem).SubItems(1) = Format(m_oCurrentVeh.Components(msCurrentSystem).FuelConsumption, "standard") & " gph"
                    'lstvSystem.ListItems.item(msCurrentSystem).SubItems(2) = Format(m_oCurrentVeh.Components(msCurrentSystem).Endurance, "standard") & " hrs"
                    '
                    ''load the column headers
                    'lstvSystem.ColumnHeaders.Add 1, , "System", 110
                    'lstvSystem.ColumnHeaders.Add 2, , "Consumption", 65
                    'lstvSystem.ColumnHeaders.Add 3, , "Endurance", 65
                    '
                    'lstvLocal.ColumnHeaders.Add 1, , "Component"
                    'lstvLocal.ColumnHeaders.Add 2, , "Capcty", 55
                    'lstvLocal.ColumnHeaders.Add 3, , "Fuel", 55
                    'lstvGlobal.ColumnHeaders.Add 1, , "Component"
                    'lstvGlobal.ColumnHeaders.Add 2, , "Capcty", 55
                    'lstvGlobal.ColumnHeaders.Add 3, , "Fuel", 55
                    '
                    'End Sub
                    '
                    'Private Sub DisplayPower()
                    'Dim consumed As Single
                    'Dim generated As Single
                    'Dim item As ListItem
                    'Dim i As Long
                    '
                    'If msCurrentPowerSystem = "" Then Exit Sub
                    '
                    ''get the generated power
                    'generated = m_oCurrentVeh.Components(msCurrentPowerSystem).Output
                    '
                    ''get the amount consumed
                    'If lstvLocal.ListItems.Count <> 0 Then
                    '    consumed = 0
                    '    For i = 1 To lstvLocal.ListItems.Count
                    '        consumed = consumed + m_oCurrentVeh.Components(lstvLocal.ListItems(i).Key).PowerReqt
                    '    Next
                    'Else
                    '    consumed = 0
                    'End If
                    '
                    ''round the value for consumed
                    'lstvSystem.ListItems.item(msCurrentPowerSystem).SubItems(2) = Format(generated - consumed, "standard")
                    'End Sub

Public Sub ShowLinks(ByVal lngLinkType As Long)


'todo: Perhaps pass in the arrays and array types so that we can use same code for Power, Fuel and Weapons?

Dim arrKeys() As String
Dim i As Long


'0) A PowerProfile node MUST BE SELECTED in main GVD treeview. If none, prompt user to create one.

'1) treeLinks - shows all SUPPLIERS
    'a) start with list of all Suppliers
    'b) given Profile, determine which is not in an existing Group
    'c) if not in an existing group, create new group and add it as sole parent
    
    
'2) lstviewLinks - shows all UNASSIGNED consumers
    'a) start with list of all Consumers
    'b) given current Profile, search all LinkGroups and return TRUE if Assigned
    'c) if not assigned, add to lstviewLinks
    'd) if assigned do nothing, the other code will already have it showing in the TreeView.
    'e)*** In fact, if we use the same key values for nodes in the treeview,  it becomes trivial to see if a node is already in the tree!!
    
frmDesigner.lstviewLinks.ListItems.Clear
frmDesigner.lstviewLinks.View = lvwReport

'get the unassigned list from the Profile

If m_oCurrentVeh.ActiveProfile = "" Then Exit Sub

Debug.Assert lngLinkType > 0


arrKeys = m_oCurrentVeh.Profiles(m_oCurrentVeh.ActiveProfile).UnAssignedConsumerList
frmDesigner.lstviewLinks.ListItems.Clear

If (UBound(arrKeys) = 0) Then
    Debug.Print "modGUI_Power:ShowLinks -- Exiting with keys = 0"
    Exit Sub
ElseIf arrKeys(1) = "" Then
    Debug.Print "modGUI_Power:ShowLinks -- Exiting with key 1 = """ & " Upper Bound = " & UBound(arrKeys)
    Exit Sub
Else
    For i = 1 To UBound(arrKeys)
        Debug.Print "modGUI_Power:ShowLinks -- Adding " & m_oCurrentVeh.Components(arrKeys(i)).CustomDescription
        frmDesigner.lstviewLinks.ListItems.Add , arrKeys(i), m_oCurrentVeh.Components(arrKeys(i)).CustomDescription
        If lngLinkType = FUEL_PROFILE Then
            frmDesigner.lstviewLinks.ListItems.Item(i).SubItems(1) = m_oCurrentVeh.Components(arrKeys(i)).Endurance & " kW"
        Else
            frmDesigner.lstviewLinks.ListItems.Item(i).SubItems(1) = m_oCurrentVeh.Components(arrKeys(i)).PowerReqt & " kW"
        End If
        frmDesigner.lstviewLinks.ListItems.Item(i).TAG = arrKeys(i)
    Next
End If



'On Error GoTo errorhandler
'if our previously default power system was deleted from the Vehicle
'set the default power system to empty
'If lstvSystem.ListItems.Count > 0 Then
'    For i = 1 To frmDesigner.lstviewLinks.ListItems.Count
'         If frmDesigner.lstviewLinks.ListItems.item(i).Key = msCurrentPowerSystem Then
'             bCurrentFound = True
'             Exit For
'        End If
'    Next
'End If

'If bCurrentFound = False Then msCurrentPowerSystem = ""

'make sure we have a Power System selected by default if none is already
'If msCurrentPowerSystem = "" And frmDesigner.lstviewLinks.ListItems.Count > 0 Then
'    msCurrentPowerSystem = frmDesigner.lstviewLinks.ListItems.item(1).Key
'    frmDesigner.lstviewLinks.ListItems.item(1).Selected = True 'give it the highlight
'ElseIf msCurrentPowerSystem = "" Then
'Else 'make sure the current power system is highlighted in the System List
'    lstvSystem.ListItems.item(msCurrentPowerSystem).Selected = True
'End If
'
''populate the local listview with the Power consuming components that are already added to
''the current power system
'If msCurrentPowerSystem <> "" Then
'    arrKeys = m_oCurrentVeh.Components(msCurrentPowerSystem).GetCurrentConsumptionSystemKeys
'    If arrKeys(1) = "" Then
'    Else
'        For i = 1 To UBound(arrKeys)
'            lstvLocal.ListItems.Add , arrKeys(i), m_oCurrentVeh.Components(arrKeys(i)).customdescription
'            lstvLocal.ListItems.item(i).SubItems(1) = m_oCurrentVeh.Components(arrKeys(i)).PowerReqt
'        Next
'    End If
'End If
'
''add any component that use power to the global component list EXCEPT for any that are already
''in the keychain of ANY Power System
'PowerConsumptionKeys = m_oCurrentVeh.Components(BODY_KEY).GetCurrentPowerConsumptionKeys
'If PowerConsumptionKeys(1) = "" Then
'    'there are no power consuming components to add
'Else
'    For i = 1 To UBound(PowerConsumptionKeys)
'        KeyFound = False 'reset the flag variable at start of outermost loop
'        For j = 1 To lstvSystem.ListItems.Count
'            If KeyFound Then Exit For
'            TempKey = lstvSystem.ListItems(j).Key
'            arrKeys = m_oCurrentVeh.Components(TempKey).GetCurrentConsumptionSystemKeys
'            For k = 1 To UBound(arrKeys)
'                If PowerConsumptionKeys(i) = arrKeys(k) Then 'this component is already assigned to another power system.  we cant use it
'                    KeyFound = True
'                    Exit For
'                End If
'            Next
'        Next
'        If KeyFound <> True Then 'the component is not already assigned so we can add it to the available global list
'            lstvGlobal.ListItems.Add , m_oCurrentVeh.Components(PowerConsumptionKeys(i)).Key, m_oCurrentVeh.Components(PowerConsumptionKeys(i)).customdescription
'            lstvGlobal.ListItems.item(lstvGlobal.ListItems.Count).SubItems(1) = m_oCurrentVeh.Components(PowerConsumptionKeys(i)).PowerReqt
'        End If
'    Next
'End If

'update the power indicator labels

   ' DisplayPower

End Sub
