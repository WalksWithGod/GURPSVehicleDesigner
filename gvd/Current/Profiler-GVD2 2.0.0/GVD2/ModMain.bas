Attribute VB_Name = "ModMain"
Option Explicit

Public Sub setGUID()
    ' if its an old GVD version veh file with no GUID, create one here MPJ 07/25/2000
vbwProfiler.vbwProcIn 1
vbwProfiler.vbwExecuteLine 1
    If (p_sGUID = Space$(39)) Or (p_sGUID = "") Then
vbwProfiler.vbwExecuteLine 2
        p_sGUID = CreateGUID  'MPJ 07/25/2000
    End If
vbwProfiler.vbwExecuteLine 3 'B
vbwProfiler.vbwProcOut 1
vbwProfiler.vbwExecuteLine 4
End Sub

Sub SetRegisteredToolbarButtonStates()
'//this determines if the program is regg'ed or not and will accordingly gray
'the buttons for Print and Saving
vbwProfiler.vbwProcIn 2
Dim tempbyte() As Byte
Dim bFlag As Boolean
Dim i As Long

vbwProfiler.vbwExecuteLine 5
On Error GoTo errorhandler

    #If DEBUG_MODE = False Then
        tempbyte = ChopCheck
        If (IsEmpty(tempbyte) = False) And (IsEmpty(gsRegNum) = False) And (UBound(tempbyte) - LBound(tempbyte) = UBound(gsRegNum) - LBound(gsRegNum)) Then
            For i = 1 To UBound(gsRegNum)
                If tempbyte(i) = gsRegNum(i) Then
                    bFlag = True
                Else
                    bFlag = False
                    Exit For
                End If
            Next
        Else
            bFlag = False
        End If
    #Else
vbwProfiler.vbwExecuteLine 6
        bFlag = True
    #End If


vbwProfiler.vbwExecuteLine 7
    If bFlag Then
        'Set the button states for the Toolbar1
        'mnuPerformance.Enabled = True
vbwProfiler.vbwExecuteLine 8
        frmDesigner.mnuSave.Enabled = True
vbwProfiler.vbwExecuteLine 9
        frmDesigner.mnuSaveAs.Enabled = True
vbwProfiler.vbwExecuteLine 10
        frmDesigner.mnuPrint.Enabled = True
vbwProfiler.vbwExecuteLine 11
        frmDesigner.mnuExport.Enabled = True
vbwProfiler.vbwExecuteLine 12
        frmDesigner.mnuPublish.Enabled = True
vbwProfiler.vbwExecuteLine 13
        With frmDesigner.Toolbar1.Buttons
vbwProfiler.vbwExecuteLine 14
            .Item(3).Enabled = True ' save button
vbwProfiler.vbwExecuteLine 15
            .Item(5).Enabled = True ' print preview
vbwProfiler.vbwExecuteLine 16
            .Item(9).Enabled = True ' publish
vbwProfiler.vbwExecuteLine 17
        End With
    Else
vbwProfiler.vbwExecuteLine 18 'B
vbwProfiler.vbwExecuteLine 19
        frmDesigner.mnuSave.Enabled = False
vbwProfiler.vbwExecuteLine 20
        frmDesigner.mnuSaveAs.Enabled = False
vbwProfiler.vbwExecuteLine 21
        frmDesigner.mnuPrint.Enabled = False
vbwProfiler.vbwExecuteLine 22
        frmDesigner.mnuExport.Enabled = False
vbwProfiler.vbwExecuteLine 23
        frmDesigner.mnuPublish.Enabled = False
vbwProfiler.vbwExecuteLine 24
        With frmDesigner.Toolbar1.Buttons
vbwProfiler.vbwExecuteLine 25
            .Item(3).Enabled = False ' save button
vbwProfiler.vbwExecuteLine 26
            .Item(5).Enabled = False ' print preview
vbwProfiler.vbwExecuteLine 27
            .Item(9).Enabled = False ' publish
vbwProfiler.vbwExecuteLine 28
        End With
    End If
vbwProfiler.vbwExecuteLine 29 'B
vbwProfiler.vbwProcOut 2
vbwProfiler.vbwExecuteLine 30
Exit Sub
errorhandler:
vbwProfiler.vbwExecuteLine 31
    bFlag = False
vbwProfiler.vbwExecuteLine 32
    frmDesigner.mnuSave.Enabled = False
vbwProfiler.vbwExecuteLine 33
    frmDesigner.mnuSaveAs.Enabled = False
vbwProfiler.vbwExecuteLine 34
    frmDesigner.mnuExport.Enabled = False
vbwProfiler.vbwExecuteLine 35
    frmDesigner.mnuPublish.Enabled = False
vbwProfiler.vbwExecuteLine 36
    With frmDesigner.Toolbar1.Buttons
vbwProfiler.vbwExecuteLine 37
        .Item(3).Enabled = False ' save button
vbwProfiler.vbwExecuteLine 38
        .Item(5).Enabled = False ' print preview
vbwProfiler.vbwExecuteLine 39
        .Item(9).Enabled = False ' publish
vbwProfiler.vbwExecuteLine 40
    End With
vbwProfiler.vbwExecuteLine 41
    ReDim gsRegName(1)
vbwProfiler.vbwExecuteLine 42
    ReDim gsRegNum(1)
vbwProfiler.vbwExecuteLine 43
    gsRegID = Empty
vbwProfiler.vbwProcOut 2
vbwProfiler.vbwExecuteLine 44
End Sub


Sub AddDependantObjects(ByVal sParentDatatype As Integer, ByVal sParentKey As String)
''Some objects require that a complimenting component gets added
vbwProfiler.vbwProcIn 3
Dim CurrentKey As String
Dim legarray() As String
Dim MirrorParentKey As String
Dim MirrorSiblingKey As String
Dim sParentParent As String
Dim lngDatatype As Long


vbwProfiler.vbwExecuteLine 45
sParentParent = frmDesigner.treeVehicle.Nodes(sParentKey).Parent.Key
vbwProfiler.vbwExecuteLine 46
CurrentKey = GetNextKey 'get a new key


vbwProfiler.vbwExecuteLine 47
If sParentDatatype = Leg Then
    'add the drivetrain for the first leg
    'now add the drivetrain for compliment leg
vbwProfiler.vbwExecuteLine 48
    With frmDesigner.treeVehicle
vbwProfiler.vbwExecuteLine 49
        .Nodes.Add sParentKey, tvwChild, CurrentKey, "leg motor", 110
        '.Nodes.item(CurrentKey).EnsureVisible ' expand the tree branch
vbwProfiler.vbwExecuteLine 50
    End With
vbwProfiler.vbwExecuteLine 51
    m_oCurrentVeh.addObject LegDrivetrain, CurrentKey, sParentKey, 110, "leg motor", False

    'determine if we need to add the complimentary leg
vbwProfiler.vbwExecuteLine 52
    legarray = m_oCurrentVeh.keymanager.GetCurrentLegKeys
vbwProfiler.vbwExecuteLine 53
    If UBound(legarray) = 1 And legarray(1) <> "" Then
        'get another new key
vbwProfiler.vbwExecuteLine 54
        CurrentKey = GetNextKey
vbwProfiler.vbwExecuteLine 55
        frmDesigner.treeVehicle.Nodes.Add sParentParent, tvwChild, CurrentKey, "leg", 12
vbwProfiler.vbwExecuteLine 56
        m_oCurrentVeh.addObject Leg, CurrentKey, sParentParent, 12, "leg", False

        'now add the drivetrain for this new leg
vbwProfiler.vbwExecuteLine 57
        sParentKey = CurrentKey
vbwProfiler.vbwExecuteLine 58
        CurrentKey = GetNextKey
vbwProfiler.vbwExecuteLine 59
        With frmDesigner.treeVehicle
vbwProfiler.vbwExecuteLine 60
            .Nodes.Add sParentKey, tvwChild, CurrentKey, "leg motor", 110
            '.Nodes.item(CurrentKey).EnsureVisible ' expand the tree branch
vbwProfiler.vbwExecuteLine 61
        End With
vbwProfiler.vbwExecuteLine 62
        m_oCurrentVeh.addObject LegDrivetrain, CurrentKey, sParentKey, 110, "leg motor", False
    End If
vbwProfiler.vbwExecuteLine 63 'B
'vbwLine 64:ElseIf p_lngDataType = AstronomicalInstruments Then
ElseIf vbwProfiler.vbwExecuteLine(64) Or p_lngDataType = AstronomicalInstruments Then
vbwProfiler.vbwExecuteLine 65
    With frmDesigner.treeVehicle
vbwProfiler.vbwExecuteLine 66
        .Nodes.Add sParentKey, tvwChild, CurrentKey, "full stabilization gear", 140
        '.Nodes.item(CurrentKey).EnsureVisible ' expand the tree branch
vbwProfiler.vbwExecuteLine 67
    End With
vbwProfiler.vbwExecuteLine 68
    m_oCurrentVeh.addObject FullStabilizationGear, CurrentKey, sParentKey, 140, "full stabilziation gear", False

'vbwLine 69:ElseIf p_lngDataType = Arm Then
ElseIf vbwProfiler.vbwExecuteLine(69) Or p_lngDataType = Arm Then
vbwProfiler.vbwExecuteLine 70
    With frmDesigner.treeVehicle
vbwProfiler.vbwExecuteLine 71
        .Nodes.Add sParentKey, tvwChild, CurrentKey, "arm motor", 66
        '.Nodes.item(CurrentKey).EnsureVisible ' expand the tree branch
vbwProfiler.vbwExecuteLine 72
    End With
vbwProfiler.vbwExecuteLine 73
    m_oCurrentVeh.addObject ArmMotor, CurrentKey, sParentKey, 66, "arm motor", False

'vbwLine 74:ElseIf p_lngDataType = Wing Then
ElseIf vbwProfiler.vbwExecuteLine(74) Or p_lngDataType = Wing Then
vbwProfiler.vbwExecuteLine 75
    frmDesigner.treeVehicle.Nodes.Add sParentParent, tvwChild, CurrentKey, "wing", 18
vbwProfiler.vbwExecuteLine 76
    m_oCurrentVeh.addObject Wing, CurrentKey, sParentParent, 18, "wing", False
    ' save all the settings
vbwProfiler.vbwExecuteLine 77
    With m_oCurrentVeh.Components(sParentKey)
vbwProfiler.vbwExecuteLine 78
        .SiblingKey = CurrentKey
vbwProfiler.vbwExecuteLine 79
        .Orientation = "right"
vbwProfiler.vbwExecuteLine 80
    End With
vbwProfiler.vbwExecuteLine 81
    With m_oCurrentVeh.Components(CurrentKey)
vbwProfiler.vbwExecuteLine 82
        .SiblingKey = sParentKey
vbwProfiler.vbwExecuteLine 83
        .Orientation = "left"
vbwProfiler.vbwExecuteLine 84
    End With
'vbwLine 85:ElseIf p_lngDataType = OrnithopterDrivetrain Then
ElseIf vbwProfiler.vbwExecuteLine(85) Or p_lngDataType = OrnithopterDrivetrain Then

    'determine the location to add the drivetrain
vbwProfiler.vbwExecuteLine 86
    With Vehicle
vbwProfiler.vbwExecuteLine 87
        MirrorParentKey = .Components(sParentKey).Parent
vbwProfiler.vbwExecuteLine 88
        MirrorSiblingKey = .Components(MirrorParentKey).SiblingKey
vbwProfiler.vbwExecuteLine 89
        frmDesigner.treeVehicle.Nodes.Add MirrorSiblingKey, tvwChild, CurrentKey, "ornithopter drivetrain", 110
vbwProfiler.vbwExecuteLine 90
        .addObject OrnithopterDrivetrain, CurrentKey, MirrorSiblingKey, 110, "ornithopter drivetrain", False
vbwProfiler.vbwExecuteLine 91
    End With
    ' save all the settings
vbwProfiler.vbwExecuteLine 92
    With m_oCurrentVeh.Components(CurrentKey)
vbwProfiler.vbwExecuteLine 93
        .SiblingKey = sParentKey
vbwProfiler.vbwExecuteLine 94
    End With
    'save the sibling key for the other wing in the pair
vbwProfiler.vbwExecuteLine 95
    m_oCurrentVeh.Components(sParentKey).SiblingKey = CurrentKey
End If
vbwProfiler.vbwExecuteLine 96 'B
vbwProfiler.vbwProcOut 3
vbwProfiler.vbwExecuteLine 97
End Sub


Function IsValidKeyCode(ByVal iKeyAscii As Integer) As Long
'    If (KeyAscii = 124) Or (KeyAscii = 64) Or (KeyAscii = 91) Or (KeyAscii = 93) Then
'        KeyAscii = 0
'    End If
vbwProfiler.vbwProcIn 4

    'note: UserEmail and the Publish textboxes are the only ones that will allow the character code 64 which is the @ symbol
    ' todo: only code which uses this is the publish email and the frmNotes.   Do i want to merge this
    ' code somewhat with the code which checks valid filenames?  I could break that up into seperte functions
    ' since part of it does check for legal characters in filenames.
vbwProfiler.vbwExecuteLine 98
    If (iKeyAscii = 124) Or (iKeyAscii = 91) Or (iKeyAscii = 93) Then
vbwProfiler.vbwExecuteLine 99
        IsValidKeyCode = False
    Else
vbwProfiler.vbwExecuteLine 100 'B
vbwProfiler.vbwExecuteLine 101
        IsValidKeyCode = True
    End If
vbwProfiler.vbwExecuteLine 102 'B
vbwProfiler.vbwProcOut 4
vbwProfiler.vbwExecuteLine 103
End Function


Function GetNextKey() As String 'todo: obsolete.  Our nodes now use their object handles for keys?
'''Returns a new key value for each Node being added to the TreeView
'''This algorithm is very simple and will limit you to adding a total of 999 nodes
'''Each node needs a unique key and to allow removing nodes you can't use the Nodes count +1
'''as the key for a new node.
''    Dim sNewKey As String
''    Dim iHold As Integer
''    Dim i As Integer
''    On Error GoTo myerr
''    'The next line will return error #35600 if there are no Nodes in the TreeView
''    iHold = Val(frmDesigner.treeVehicle.Nodes(1).Key)
''    For i = 1 To frmDesigner.treeVehicle.Nodes.Count
''        If Val(frmDesigner.treeVehicle.Nodes(i).Key) > iHold Then
''            iHold = Val(frmDesigner.treeVehicle.Nodes(i).Key)
''        End If
''    Next
''    iHold = iHold + 1
''    sNewKey = KeyFromLong(iHold)
''    GetNextKey = sNewKey 'Return a unique key
''    Exit Function
''myerr:
''    'Because the TreeView is empty return a 1 for the key of the first Node
''    Debug.Print "ModMain.GetNextKey - Error # " & err.Number & " " & err.Description & " This is a handled error.  No problem here."
''    GetNextKey = "1_"
''    Exit Function
vbwProfiler.vbwProcIn 5
vbwProfiler.vbwProcOut 5
vbwProfiler.vbwExecuteLine 104
End Function



Sub RemoveComponent(ByVal Key As String)
vbwProfiler.vbwProcIn 6
Dim SiblingKey As String
Dim legkey() As String

vbwProfiler.vbwExecuteLine 105
legkey = m_oCurrentVeh.keymanager.GetCurrentLegKeys
    'Remove the selected Node and the Collection item

    'On Error GoTo myerr 'if the treeview does not have a node selected
    ' the next line of code will return an error number 91
    'iIndex = frmDesigner.treeVehicle.SelectedItem.Index 'Check to see if a Node is selected

vbwProfiler.vbwExecuteLine 106
If DeleteCheck(Key) Then 'make sure the user is allowed to delte this object

    'if its a wing, then delete the wing pair
    'IMPORTANT: Note that this type of stuff is purely User Ease of Use type stuff.
    'Realisticly, if a person wanted to create a contraption that only had one wing or one
    'leg, why not?  Or even in battle, one wing can get blown off... perhaps the plane is on
    ' the ground and a bomb raid destroys a wing on a plane still on the ground.  Point is,
    ' this stuff is non in vehicles.dll for a reason, its strictly UI fluff and not core
    ' vehicles rules or statistics.
vbwProfiler.vbwExecuteLine 107
    If m_oCurrentVeh.Components(Key).Datatype = Wing Then
         'If the Node has Children call the sub that deletes the children
vbwProfiler.vbwExecuteLine 108
        If frmDesigner.treeVehicle.Nodes(Key).children > 0 Then
vbwProfiler.vbwExecuteLine 109
            RemoveChild frmDesigner.treeVehicle.Nodes(Key).index
        End If
vbwProfiler.vbwExecuteLine 110 'B
vbwProfiler.vbwExecuteLine 111
        SiblingKey = m_oCurrentVeh.Components(Key).SiblingKey
vbwProfiler.vbwExecuteLine 112
        frmDesigner.treeVehicle.Nodes.Remove Key 'Removes the Node and any children it has

vbwProfiler.vbwExecuteLine 113
        m_oCurrentVeh.keymanager.RemoveSubAssemblyKey Key
vbwProfiler.vbwExecuteLine 114
        m_oCurrentVeh.Components.Remove Key ' remove the item from the collection
         'If the Sibling has Children call the sub that deletes the children
vbwProfiler.vbwExecuteLine 115
        If SiblingKey <> "" Then
vbwProfiler.vbwExecuteLine 116
            If frmDesigner.treeVehicle.Nodes(SiblingKey).children > 0 Then
vbwProfiler.vbwExecuteLine 117
                RemoveChild frmDesigner.treeVehicle.Nodes(SiblingKey).index
            End If
vbwProfiler.vbwExecuteLine 118 'B
vbwProfiler.vbwExecuteLine 119
            frmDesigner.treeVehicle.Nodes.Remove SiblingKey 'Removes the Node and any children it has
vbwProfiler.vbwExecuteLine 120
            m_oCurrentVeh.Components.Remove SiblingKey ' remove the item from the collection
vbwProfiler.vbwExecuteLine 121
            m_oCurrentVeh.keymanager.RemoveSubAssemblyKey SiblingKey
        End If
vbwProfiler.vbwExecuteLine 122 'B
    'if its an ornithoper , delete its sibling 'todo: why the fuck would this drivetrain have siblings anyway?
'vbwLine 123:    ElseIf m_oCurrentVeh.Components(Key).Datatype = OrnithopterDrivetrain Then
    ElseIf vbwProfiler.vbwExecuteLine(123) Or m_oCurrentVeh.Components(Key).Datatype = OrnithopterDrivetrain Then
         'If the Node has Children call the sub that deletes the children
vbwProfiler.vbwExecuteLine 124
        If frmDesigner.treeVehicle.Nodes(Key).children > 0 Then
vbwProfiler.vbwExecuteLine 125
            RemoveChild frmDesigner.treeVehicle.Nodes(Key).index
        End If
vbwProfiler.vbwExecuteLine 126 'B
vbwProfiler.vbwExecuteLine 127
        SiblingKey = m_oCurrentVeh.Components(Key).SiblingKey
vbwProfiler.vbwExecuteLine 128
        frmDesigner.treeVehicle.Nodes.Remove Key 'Removes the Node and any children it has
vbwProfiler.vbwExecuteLine 129
        m_oCurrentVeh.keymanager.RemoveKeyChainKey Key, m_oCurrentVeh.Components(Key).Datatype 'remove any keychain references
vbwProfiler.vbwExecuteLine 130
        m_oCurrentVeh.Components.Remove Key ' remove the item from the collection
         'If the Sibling has Children call the sub that deletes the children
vbwProfiler.vbwExecuteLine 131
        If frmDesigner.treeVehicle.Nodes(SiblingKey).children > 0 Then
vbwProfiler.vbwExecuteLine 132
            RemoveChild frmDesigner.treeVehicle.Nodes(SiblingKey).index
        End If
vbwProfiler.vbwExecuteLine 133 'B
vbwProfiler.vbwExecuteLine 134
        frmDesigner.treeVehicle.Nodes.Remove SiblingKey 'Removes the Node and any children it has
vbwProfiler.vbwExecuteLine 135
        m_oCurrentVeh.keymanager.RemoveKeyChainKey SiblingKey, m_oCurrentVeh.Components(SiblingKey).Datatype 'remove any keychain references
vbwProfiler.vbwExecuteLine 136
        m_oCurrentVeh.Components.Remove SiblingKey ' remove the item from the collection
    'if its a leg and the Leg array has ONLY 2 legs, then delete the leg pair
'vbwLine 137:    ElseIf m_oCurrentVeh.Components(Key).Datatype = Leg And UBound(legkey) = 2 Then
    ElseIf vbwProfiler.vbwExecuteLine(137) Or m_oCurrentVeh.Components(Key).Datatype = Leg And UBound(legkey) = 2 Then
          'If the Node has Children call the sub that deletes the children
vbwProfiler.vbwExecuteLine 138
        If frmDesigner.treeVehicle.Nodes(legkey(1)).children > 0 Then
vbwProfiler.vbwExecuteLine 139
            RemoveChild frmDesigner.treeVehicle.Nodes(legkey(1)).index
        End If
vbwProfiler.vbwExecuteLine 140 'B
vbwProfiler.vbwExecuteLine 141
        frmDesigner.treeVehicle.Nodes.Remove legkey(1) 'Removes the Node and any children it has
vbwProfiler.vbwExecuteLine 142
        m_oCurrentVeh.keymanager.RemoveLegKey legkey(1)
vbwProfiler.vbwExecuteLine 143
        m_oCurrentVeh.keymanager.RemoveSubAssemblyKey legkey(1) 'MPJ 07/25/2000 was using index 2 instead of 1 for the legkey()
vbwProfiler.vbwExecuteLine 144
        m_oCurrentVeh.Components.Remove legkey(1) ' remove the item from the collection
         'If the Sibling has Children call the sub that deletes the children
vbwProfiler.vbwExecuteLine 145
        If frmDesigner.treeVehicle.Nodes(legkey(2)).children > 0 Then
vbwProfiler.vbwExecuteLine 146
            RemoveChild frmDesigner.treeVehicle.Nodes(legkey(2)).index
        End If
vbwProfiler.vbwExecuteLine 147 'B
vbwProfiler.vbwExecuteLine 148
        frmDesigner.treeVehicle.Nodes.Remove legkey(2) 'Removes the Node and any children it has
vbwProfiler.vbwExecuteLine 149
        m_oCurrentVeh.keymanager.RemoveLegKey legkey(2)
vbwProfiler.vbwExecuteLine 150
        m_oCurrentVeh.keymanager.RemoveSubAssemblyKey legkey(2)
vbwProfiler.vbwExecuteLine 151
        m_oCurrentVeh.Components.Remove legkey(2) ' remove the item from the collection

    'do normal delete of component
    Else
vbwProfiler.vbwExecuteLine 152 'B
        'If the Node has Children call the sub that deletes the children
vbwProfiler.vbwExecuteLine 153
        If frmDesigner.treeVehicle.Nodes(Key).children > 0 Then
vbwProfiler.vbwExecuteLine 154
            RemoveChild frmDesigner.treeVehicle.Nodes(Key).index
        End If
vbwProfiler.vbwExecuteLine 155 'B
vbwProfiler.vbwExecuteLine 156
        frmDesigner.treeVehicle.Nodes.Remove Key 'Removes the Node and any children it has
vbwProfiler.vbwExecuteLine 157
        m_oCurrentVeh.keymanager.RemoveKeyChainKey Key, m_oCurrentVeh.Components(Key).Datatype 'remove any keychain references
vbwProfiler.vbwExecuteLine 158
        m_oCurrentVeh.Components.Remove Key ' remove the item from the collection

    End If
vbwProfiler.vbwExecuteLine 159 'B
End If
vbwProfiler.vbwExecuteLine 160 'B


    'Exit Sub
'myerr:
    'Display a messge telling the user to select a node
 '   MsgBox ("Nothing to delete")
    'Exit Sub
vbwProfiler.vbwProcOut 6
vbwProfiler.vbwExecuteLine 161
End Sub
    
Function DeleteCheck(Key As String) As Boolean
    'NOTE: In terms of MMORPG anti cheat security, i dont think it matters
    ' that DeleteChecks are performed on the GUI side (which translates to Client side).
    ' Deleting a componet would diminish a vehicle right?  Only problem i could forsee is if
    ' there was a component that the user would not want on his vehicle for some reason but
    ' was forced to have it... perhaps because it prevented him from buying something else
    ' or perhaps because it adds weight/volume etc and he'd rather remove it and replace it with
    ' something better or just leave it empty
vbwProfiler.vbwProcIn 7

    ' maybe this isnt such a big deal..  this function does seem very GUI related and not
    ' really important to the vehicle.dll much at all.

    Dim lngDatatype As Integer

vbwProfiler.vbwExecuteLine 162
    lngDatatype = m_oCurrentVeh.Components(Key).Datatype

    'Delete checks for arms motors
vbwProfiler.vbwExecuteLine 163
    If lngDatatype = ArmMotor Then
vbwProfiler.vbwExecuteLine 164
        MsgBox "The Arm Motor is required by the Arm Assembly and cannot be deleted."
vbwProfiler.vbwExecuteLine 165
        DeleteCheck = False
vbwProfiler.vbwProcOut 7
vbwProfiler.vbwExecuteLine 166
        Exit Function
    'delete check for leg drivetrains
'vbwLine 167:    ElseIf lngDatatype = LegDrivetrain Then
    ElseIf vbwProfiler.vbwExecuteLine(167) Or lngDatatype = LegDrivetrain Then
vbwProfiler.vbwExecuteLine 168
        MsgBox "The Leg Motor is required by the Leg Assembly and cannot be deleted."
vbwProfiler.vbwExecuteLine 169
        DeleteCheck = False
vbwProfiler.vbwProcOut 7
vbwProfiler.vbwExecuteLine 170
        Exit Function
'vbwLine 171:    ElseIf lngDatatype = Leg And UBound(m_oCurrentVeh.keymanager.GetCurrentLegKeys) = 2 Then
    ElseIf vbwProfiler.vbwExecuteLine(171) Or lngDatatype = Leg And UBound(m_oCurrentVeh.keymanager.GetCurrentLegKeys) = 2 Then
vbwProfiler.vbwExecuteLine 172
        If MsgBox("The last two legs on a Vehicle are deleted in pairs. Do you wish to delete the remaining two legs?", vbYesNo, "Delete remaining legs?") = vbNo Then
vbwProfiler.vbwExecuteLine 173
            DeleteCheck = False
vbwProfiler.vbwProcOut 7
vbwProfiler.vbwExecuteLine 174
            Exit Function
        End If
vbwProfiler.vbwExecuteLine 175 'B
    'warn the user that wings are deleted in pairs
'vbwLine 176:    ElseIf lngDatatype = Wing Then
    ElseIf vbwProfiler.vbwExecuteLine(176) Or lngDatatype = Wing Then
vbwProfiler.vbwExecuteLine 177
        If MsgBox("Wings are deleted in pairs. Do you wish to delete the Wing pair?", vbYesNo, "Delete Wing Pair?") = vbNo Then
vbwProfiler.vbwExecuteLine 178
            DeleteCheck = False
vbwProfiler.vbwProcOut 7
vbwProfiler.vbwExecuteLine 179
            Exit Function
        End If
vbwProfiler.vbwExecuteLine 180 'B
    'full stabilization gear from Astronimical Instruments
'vbwLine 181:    ElseIf lngDatatype = FullStabilizationGear And m_oCurrentVeh.Components(m_oCurrentVeh.Components(Key).LogicalParent).Datatype = AstronomicalInstruments Then
    ElseIf vbwProfiler.vbwExecuteLine(181) Or lngDatatype = FullStabilizationGear And m_oCurrentVeh.Components(m_oCurrentVeh.Components(Key).LogicalParent).Datatype = AstronomicalInstruments Then
vbwProfiler.vbwExecuteLine 182
        MsgBox "Full Stabilization Gear is required for Astronomical Instruments and cannot be deleted."
vbwProfiler.vbwExecuteLine 183
        DeleteCheck = False
vbwProfiler.vbwProcOut 7
vbwProfiler.vbwExecuteLine 184
        Exit Function
    End If
vbwProfiler.vbwExecuteLine 185 'B

    'Delete checks for form fitting battlesuit
vbwProfiler.vbwExecuteLine 186
    If m_oCurrentVeh.Options.BattleSuit = "Form Fitting" Then
vbwProfiler.vbwExecuteLine 187
        Select Case Key
'vbwLine 188:        Case "2_", "3_", "4_", "6_", "8_", "9_", "10_", "11_"
        Case IIf(vbwProfiler.vbwExecuteLine(188), VBWPROFILER_EMPTY, _
        "2_"), "3_", "4_", "6_", "8_", "9_", "10_", "11_"
vbwProfiler.vbwExecuteLine 189
            MsgBox "This component is required by the Battlesuit and cannot be deleted."
vbwProfiler.vbwExecuteLine 190
            DeleteCheck = False
vbwProfiler.vbwProcOut 7
vbwProfiler.vbwExecuteLine 191
            Exit Function
        End Select
vbwProfiler.vbwExecuteLine 192 'B
'vbwLine 193:    ElseIf m_oCurrentVeh.Options.BattleSuit = "Pilot in Body" Then
    ElseIf vbwProfiler.vbwExecuteLine(193) Or m_oCurrentVeh.Options.BattleSuit = "Pilot in Body" Then
vbwProfiler.vbwExecuteLine 194
        If Key = "2_" Then
vbwProfiler.vbwExecuteLine 195
            MsgBox "This component is required by the Battlesuit and cannot be deleted."
vbwProfiler.vbwExecuteLine 196
            DeleteCheck = False
vbwProfiler.vbwProcOut 7
vbwProfiler.vbwExecuteLine 197
            Exit Function
        End If
vbwProfiler.vbwExecuteLine 198 'B
'vbwLine 199:    ElseIf m_oCurrentVeh.Options.BattleSuit = "Pilot in Turret" Then
    ElseIf vbwProfiler.vbwExecuteLine(199) Or m_oCurrentVeh.Options.BattleSuit = "Pilot in Turret" Then
vbwProfiler.vbwExecuteLine 200
        If Key = "2_" Or Key = "3_" Then
vbwProfiler.vbwExecuteLine 201
            MsgBox "This component is required by the Battlesuit and cannot be deleted."
vbwProfiler.vbwExecuteLine 202
            DeleteCheck = False
vbwProfiler.vbwProcOut 7
vbwProfiler.vbwExecuteLine 203
            Exit Function
        End If
vbwProfiler.vbwExecuteLine 204 'B
    End If
vbwProfiler.vbwExecuteLine 205 'B

    'MPJ 03/30/02 -- I Believe All This is now obsolete.  Since Ground performances can be created
    'without first having a ground subassembly assigned to them.  In fact, the routines are now setup
    'to use the first ground subassembly it fines.  Now we only need to delete LEG PAIRS

    'OBSOLETE - warn the user that certain GROUND assemblies like Legs, Wheels, tracks etc if deleted
    'OBSOLETE - will result in the Performance Profile that is dependant on it to be deleted as well
    'OBSOLETE - Select Case Datatype
    '    Case Leg, Track, Wheel, Skid
    '        If MsgBox("Deleting this component will result in any Performance Profiles that are linked to it to be deleted as well.  Continue?", vbYesNo, "Delete Warning...") = vbNo Then
    '            DeleteCheck = False
    '            Exit Function
    '        Else 'warn the user that the last 2 legs are deleted in pairs.
    '            If Datatype = Leg And UBound(Veh.KeyManager.GetCurrentLegKeys) = 2 Then
    '                    If MsgBox("The last two legs on a Vehicle are deleted in pairs. Do you wish to delete the remaining two legs?", vbYesNo, "Delete remaining legs?") = vbNo Then
    '                        DeleteCheck = False
    '                        Exit Function
    '                    Else
    '                        RemoveAnyDependantPerformanceProfiles Key
    '                    End If
    '            Else
    '                RemoveAnyDependantPerformanceProfiles Key
    '            End If
    '        End If
    'End Select

vbwProfiler.vbwExecuteLine 206
    DeleteCheck = True ' <-- Note, leg
vbwProfiler.vbwProcOut 7
vbwProfiler.vbwExecuteLine 207
End Function

Sub RemoveChild(ByVal iNodeIndex As Integer)
    'todo: Obsolete -- treeX contains a single method to delete all children
    '      further, our hieararchal storage of components allows for deleting of parent
    '      with it taking care of deleting all its children for us
    'This sub uses recursion to loop through the child nodes and delete them.
    'It receives the Index of the node that has the children
vbwProfiler.vbwProcIn 8
    Dim i As Integer
    Dim iTempIndex As Integer
    Dim Key As String

vbwProfiler.vbwExecuteLine 208
    iTempIndex = frmDesigner.treeVehicle.Nodes(iNodeIndex).Child.FirstSibling.index
vbwProfiler.vbwExecuteLine 209
    Key = frmDesigner.treeVehicle.Nodes(iTempIndex).Key

    'Loop through all a Parents Child Nodes
vbwProfiler.vbwExecuteLine 210
    For i = 1 To frmDesigner.treeVehicle.Nodes(iNodeIndex).children

vbwProfiler.vbwExecuteLine 211
        m_oCurrentVeh.keymanager.RemoveKeyChainKey Key, m_oCurrentVeh.Components(Key).Datatype 'remove any keychain references
vbwProfiler.vbwExecuteLine 212
        m_oCurrentVeh.Components.Remove Key

        ' If the Node we are on has a child call the Sub again
vbwProfiler.vbwExecuteLine 213
        If frmDesigner.treeVehicle.Nodes(iTempIndex).children > 0 Then
vbwProfiler.vbwExecuteLine 214
           RemoveChild (iTempIndex) ' <--recurse
        End If
vbwProfiler.vbwExecuteLine 215 'B

      ' If we are not on the last child move to the next child Node
vbwProfiler.vbwExecuteLine 216
        If i <> frmDesigner.treeVehicle.Nodes(iNodeIndex).children Then
vbwProfiler.vbwExecuteLine 217
           iTempIndex = frmDesigner.treeVehicle.Nodes(iTempIndex).Next.index
vbwProfiler.vbwExecuteLine 218
           Key = frmDesigner.treeVehicle.Nodes(iTempIndex).Key
        End If
vbwProfiler.vbwExecuteLine 219 'B
vbwProfiler.vbwExecuteLine 220
    Next
vbwProfiler.vbwProcOut 8
vbwProfiler.vbwExecuteLine 221
End Sub


Public Function ChopCheck2() As Byte()
    'NOTE:  This is just one of the local reg number checkers.  There will be several of these so
    ' that a hacker will have to do some serious code hacking to disable them
    ' The other called "ChopCheck" is in the modTextOutput.bas.
    ' Also check the frmCredits for related code
vbwProfiler.vbwProcIn 9


    '==(MPJ 02/20/02 - This so called formula really sucks.
    'Wish I woulda spent an additional day thinking about this, instead, i wrote in
    ' about 30 mins on the day before I went "gold."  Too late to change at this point)
    ' As far as I know, there arent any cracks for it, the problem is that the
    ' reg codes it produces arent as varied as I'd like.  Typically, only several of the
    ' characters in the code are different as compared to other generated codes for other users)==

    Dim tempbyte() As Byte
    Dim i As Long
    Dim j As Single
    Dim sTName As String
    Dim lngtotal As Single
    Dim sRegNumber As String
vbwProfiler.vbwExecuteLine 222
    ReDim tempbyte(1)

vbwProfiler.vbwExecuteLine 223
    MsgBox "ModMain.ChopCheck2() -- Function temporarily disabled..."
vbwProfiler.vbwProcOut 9
vbwProfiler.vbwExecuteLine 224
    Exit Function
'    On Error GoTo errorhandler
'    'here's the reg key formula
'    '1- the user's reg name and key are accepted into a byte array with each
'    '   letter being actually the ascii code for that letter.  Total them up
'    For i = 1 To UBound(gsRegName)
'        lngtotal = lngtotal + gsRegName(i)
'        'at the same time total the ascii value for every even valued ascii code
'        If gsRegName(i) Mod 2 = 0 Then
'            lngtotal = lngtotal + gsRegName(i)
'        End If
'    Next
'    '2 - the RegID is actually just a modifier to prevent two people having the same
'    '    name winding up with the same ID.  This ID is unique and alone can be used
'    '   to identify a user.  Multiply this to the total
'    lngtotal = lngtotal * gsRegID
'    '3- take the ascii value of the typename of the Body and multiply that to it
'    sTName = TypeName(m_oCurrentVeh.Components(BODY_KEY))
'    For i = 1 To Len(sTName)
'        lngtotal = lngtotal * Asc(Mid(sTName, i, 1))
'    Next
'    '6- take a random seed to generate the seeded random number and multiply that
'    Rnd -1
'    Randomize 9921988
'    lngtotal = lngtotal * Rnd()
'    '8- return this as a byte array that we can compare with our current one
'    'how do we split this up into seperate bytes? well we know our ascii values
'    'must be between 48-57, 65-90 and 97-122
'    'well, we can generate a random reg code based on each number in the string
'    'representation using the random seed of each number
'    For i = 1 To Len(Str(lngtotal))
'        j = Rnd()
'        If j <= 0.33 Then
'            ReDim Preserve tempbyte(i)
'            Rnd -1
'            Randomize Asc(Mid(Str(lngtotal), i, 1))
'            tempbyte(i) = Int((57 - 48 + 1) * Rnd + 48)
'            sRegNumber = sRegNumber & Chr(tempbyte(i))
'        ElseIf j <= 0.66 Then
'            ReDim Preserve tempbyte(i)
'            Rnd -1
'            Randomize Asc(Mid(Str(lngtotal), i, 1))
'            tempbyte(i) = Int((90 - 65 + 1) * Rnd + 65)
'            sRegNumber = sRegNumber & Chr(tempbyte(i))
'        Else
'            ReDim Preserve tempbyte(i)
'            Rnd -1
'            Randomize Asc(Mid(Str(lngtotal), i, 1))
'            tempbyte(i) = Int((122 - 97 + 1) * Rnd + 97)
'            sRegNumber = sRegNumber & Chr(tempbyte(i))
'        End If
'    Next
'
'    ChopCheck2 = tempbyte
'    Exit Function
'errorhandler:
'    ReDim tempbyte(1)
'    ChopCheck2 = tempbyte
vbwProfiler.vbwProcOut 9
vbwProfiler.vbwExecuteLine 225
End Function



' load Category for Type of vehicle the user is creating. This is
' used primarily in the user's print output and for web site sorting of .veh's
Public Function LoadCategories(vResults() As Variant) As Long
vbwProfiler.vbwProcIn 10

    Dim nFile As Long
    Dim sFileName As String
    Dim s As String
    Dim sEntry() As String
    Dim sResults() As String
    Dim i As Long

vbwProfiler.vbwExecuteLine 226
    On Error GoTo errorhandler
vbwProfiler.vbwExecuteLine 227
    nFile = FreeFile
vbwProfiler.vbwExecuteLine 228
    sFileName = "categories.txt"

vbwProfiler.vbwExecuteLine 229
    ChDir App.Path
vbwProfiler.vbwExecuteLine 230
    Open App.Path & "\lists\" & sFileName For Input As #nFile

'vbwLine 231:    Do While Not EOF(nFile)
    Do While vbwProfiler.vbwExecuteLine(231) Or Not EOF(nFile)
vbwProfiler.vbwExecuteLine 232
         Line Input #nFile, s
vbwProfiler.vbwExecuteLine 233
         sEntry = Split(s, ",")
vbwProfiler.vbwExecuteLine 234
         ReDim Preserve vResults(i)
vbwProfiler.vbwExecuteLine 235
         vResults(i) = sEntry(0)
vbwProfiler.vbwExecuteLine 236
         i = i + 0
         'cboCategory.AddItem sEntry(0)
vbwProfiler.vbwExecuteLine 237
    Loop
vbwProfiler.vbwExecuteLine 238
    LoadCategories = True
vbwProfiler.vbwExecuteLine 239
    Close #nFile
vbwProfiler.vbwProcOut 10
vbwProfiler.vbwExecuteLine 240
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 241
    Close #nFile
vbwProfiler.vbwExecuteLine 242
    LoadCategories = False
vbwProfiler.vbwProcOut 10
vbwProfiler.vbwExecuteLine 243
End Function

' Function called by LoadCategories to
' load subcategory for Type of vehicle the user is creating. This is
' used primarily in the user's print output and for web site sorting of .veh's
Public Function LoadSubCategories(ByRef sCategory As String, vResults() As Variant) As Long
vbwProfiler.vbwProcIn 11
    Dim sEntry As String
    Dim sFirstEntry As String
    Dim sFileName As String
    Dim i As Integer
    Dim nFile As Long
    Dim s() As String


vbwProfiler.vbwExecuteLine 244
    On Error GoTo errorhandler

vbwProfiler.vbwExecuteLine 245
    DoEvents 'make sure the combo drop down repaints

    'Clear the subcategory items

vbwProfiler.vbwExecuteLine 246
    sFileName = "categories.txt"
vbwProfiler.vbwExecuteLine 247
    nFile = FreeFile

    'make sure we are back in the program's install path
vbwProfiler.vbwExecuteLine 248
    ChDir App.Path
    ' Load the combo2 with the names of the components within the selected List file
vbwProfiler.vbwExecuteLine 249
    Open App.Path & "\lists\" & sFileName For Input As #nFile  ' Open file for input.

'vbwLine 250:    Do While Not EOF(nFile) ' Loop until end of file.
    Do While vbwProfiler.vbwExecuteLine(250) Or Not EOF(nFile) ' Loop until end of file.
vbwProfiler.vbwExecuteLine 251
        Line Input #nFile, sEntry
vbwProfiler.vbwExecuteLine 252
        s = Split(sEntry, ",")
vbwProfiler.vbwExecuteLine 253
        If s(0) = sCategory Then
vbwProfiler.vbwExecuteLine 254
            ReDim vResults(UBound(s) - 1)
vbwProfiler.vbwExecuteLine 255
            For i = 0 To UBound(s) - 1
vbwProfiler.vbwExecuteLine 256
                vResults(0) = s(i + 1)
vbwProfiler.vbwExecuteLine 257
            Next
        End If
vbwProfiler.vbwExecuteLine 258 'B
vbwProfiler.vbwExecuteLine 259
    Loop
vbwProfiler.vbwExecuteLine 260
    Close #nFile    ' Close file.
vbwProfiler.vbwProcOut 11
vbwProfiler.vbwExecuteLine 261
    Exit Function
errorhandler:
vbwProfiler.vbwExecuteLine 262
    If err.Number <> 0 Then
vbwProfiler.vbwExecuteLine 263
        InfoPrint 1, "Err in LoadSubCategories: " + err.Description
    End If
vbwProfiler.vbwExecuteLine 264 'B
vbwProfiler.vbwExecuteLine 265
    Close #nFile

vbwProfiler.vbwProcOut 11
vbwProfiler.vbwExecuteLine 266
End Function


'Private Sub SelectImage()
'    Dim fsoFolder As Folder
'    Dim fsoObject As FileSystemObject
'    Dim sImagesPath As String
'    Dim i As Long
'
'    If cmdSelectImage.Caption = "Cancel Image" Then
'        cmdSelectImage.Caption = "Select image from 'images' directory"
'        fileImages.Visible = False
'        lblVehicleImageFileName = "(no image set)"
'        m_oCurrentVeh.Components(BODY_KEY).VehicleImageFileName = ""
'        imgVehicleImage.Visible = False
'    Else
'        sImagesPath = App.Path & "\images"
'        '//check if our "images" directory exists and create it if it doesnt
'        Set fsoObject = New FileSystemObject
'        If fsoObject.FolderExists(sImagesPath) = False Then
'            fsoObject.CreateFolder sImagesPath
'        End If
'        '//set our fileImages to the Images directory
'        fileImages.Path = sImagesPath
'        '//set our filter to only look for JPG's
'        fileImages.Pattern = "*.jpg"
'        fileImages.Visible = True
'        fileImages.Refresh
'        For i = 0 To fileImages.ListCount - 1
'            fileImages.Selected(i) = False
'        Next
'        cmdSelectImage.Caption = "Cancel Image"
'    End If
'
'End Sub
'
'Private Sub fileImages_Click()
'    p_bChangedFlag = True ' JAW 2000.05.07
'
'    '//Check that its the proper resolution
'    If fileImages.FileName <> "" Then
'        imgVehicleImage = LoadPicture(App.Path & "\images\" & fileImages.FileName)
'        DoEvents
'        If imgVehicleImage.ScaleHeight <> 240 Or imgVehicleImage.ScaleWidth <> 390 Then
'            imgVehicleImage.Visible = False
'            MsgBox "Image resolution invalid. Vehicle image must be 390 x 240 pixels."
'            lblVehicleImageFileName = "(no image set)"
'            m_oCurrentVeh.Components(BODY_KEY).VehicleImageFileName = ""
'        Else
'            '//set the label
'            lblVehicleImageFileName = fileImages.FileName
'            fileImages.Visible = False
'            imgVehicleImage.Visible = True
'            m_oCurrentVeh.Components(BODY_KEY).VehicleImageFileName = fileImages.FileName
'        End If
'    Else
'        lblVehicleImageFileName = "(no image set)"
'        m_oCurrentVeh.Components(BODY_KEY).VehicleImageFileName = ""
'        imgVehicleImage.Visible = False
'
'    End If
'
'End Sub



