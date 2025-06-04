Attribute VB_Name = "ModMain"
Option Explicit

Public Sub setGUID()
    ' if its an old GVD version veh file with no GUID, create one here MPJ 07/25/2000
    If (p_sGUID = Space$(39)) Or (p_sGUID = "") Then
        p_sGUID = CreateGUID  'MPJ 07/25/2000
    End If
End Sub

Sub SetRegisteredToolbarButtonStates()
'//this determines if the program is regg'ed or not and will accordingly gray
'the buttons for Print and Saving
Dim tempbyte() As Byte
Dim bFlag As Boolean
Dim i As Long

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
        bFlag = True
    #End If
        
        
    If bFlag Then
        'Set the button states for the Toolbar1
        'mnuPerformance.Enabled = True
        frmDesigner.mnuSave.Enabled = True
        frmDesigner.mnuSaveAs.Enabled = True
        frmDesigner.mnuPrint.Enabled = True
        frmDesigner.mnuExport.Enabled = True
        frmDesigner.mnuPublish.Enabled = True
        With frmDesigner.Toolbar1.Buttons
            .Item(3).Enabled = True ' save button
            .Item(5).Enabled = True ' print preview
            .Item(9).Enabled = True ' publish
        End With
    Else
        frmDesigner.mnuSave.Enabled = False
        frmDesigner.mnuSaveAs.Enabled = False
        frmDesigner.mnuPrint.Enabled = False
        frmDesigner.mnuExport.Enabled = False
        frmDesigner.mnuPublish.Enabled = False
        With frmDesigner.Toolbar1.Buttons
            .Item(3).Enabled = False ' save button
            .Item(5).Enabled = False ' print preview
            .Item(9).Enabled = False ' publish
        End With
    End If
Exit Sub
errorhandler:
    bFlag = False
    frmDesigner.mnuSave.Enabled = False
    frmDesigner.mnuSaveAs.Enabled = False
    frmDesigner.mnuExport.Enabled = False
    frmDesigner.mnuPublish.Enabled = False
    With frmDesigner.Toolbar1.Buttons
        .Item(3).Enabled = False ' save button
        .Item(5).Enabled = False ' print preview
        .Item(9).Enabled = False ' publish
    End With
    ReDim gsRegName(1)
    ReDim gsRegNum(1)
    gsRegID = Empty
End Sub


Sub AddDependantObjects(ByVal sParentDatatype As Integer, ByVal sParentKey As String)
''Some objects require that a complimenting component gets added
Dim CurrentKey As String
Dim legarray() As String
Dim MirrorParentKey As String
Dim MirrorSiblingKey As String
Dim sParentParent As String
Dim lngDatatype As Long


sParentParent = frmDesigner.treeVehicle.Nodes(sParentKey).Parent.Key
CurrentKey = GetNextKey 'get a new key


If sParentDatatype = Leg Then
    'add the drivetrain for the first leg
    'now add the drivetrain for compliment leg
    With frmDesigner.treeVehicle
        .Nodes.Add sParentKey, tvwChild, CurrentKey, "leg motor", 110
        '.Nodes.item(CurrentKey).EnsureVisible ' expand the tree branch
    End With
    m_oCurrentVeh.addObject LegDrivetrain, CurrentKey, sParentKey, 110, "leg motor", False

    'determine if we need to add the complimentary leg
    legarray = m_oCurrentVeh.keymanager.GetCurrentLegKeys
    If UBound(legarray) = 1 And legarray(1) <> "" Then
        'get another new key
        CurrentKey = GetNextKey
        frmDesigner.treeVehicle.Nodes.Add sParentParent, tvwChild, CurrentKey, "leg", 12
        m_oCurrentVeh.addObject Leg, CurrentKey, sParentParent, 12, "leg", False
        
        'now add the drivetrain for this new leg
        sParentKey = CurrentKey
        CurrentKey = GetNextKey
        With frmDesigner.treeVehicle
            .Nodes.Add sParentKey, tvwChild, CurrentKey, "leg motor", 110
            '.Nodes.item(CurrentKey).EnsureVisible ' expand the tree branch
        End With
        m_oCurrentVeh.addObject LegDrivetrain, CurrentKey, sParentKey, 110, "leg motor", False
    End If
ElseIf p_lngDataType = AstronomicalInstruments Then
    With frmDesigner.treeVehicle
        .Nodes.Add sParentKey, tvwChild, CurrentKey, "full stabilization gear", 140
        '.Nodes.item(CurrentKey).EnsureVisible ' expand the tree branch
    End With
    m_oCurrentVeh.addObject FullStabilizationGear, CurrentKey, sParentKey, 140, "full stabilziation gear", False

ElseIf p_lngDataType = Arm Then
    With frmDesigner.treeVehicle
        .Nodes.Add sParentKey, tvwChild, CurrentKey, "arm motor", 66
        '.Nodes.item(CurrentKey).EnsureVisible ' expand the tree branch
    End With
    m_oCurrentVeh.addObject ArmMotor, CurrentKey, sParentKey, 66, "arm motor", False
    
ElseIf p_lngDataType = Wing Then
    frmDesigner.treeVehicle.Nodes.Add sParentParent, tvwChild, CurrentKey, "wing", 18
    m_oCurrentVeh.addObject Wing, CurrentKey, sParentParent, 18, "wing", False
    ' save all the settings
    With m_oCurrentVeh.Components(sParentKey)
        .SiblingKey = CurrentKey
        .Orientation = "right"
    End With
    With m_oCurrentVeh.Components(CurrentKey)
        .SiblingKey = sParentKey
        .Orientation = "left"
    End With
ElseIf p_lngDataType = OrnithopterDrivetrain Then

    'determine the location to add the drivetrain
    With Vehicle
        MirrorParentKey = .Components(sParentKey).Parent
        MirrorSiblingKey = .Components(MirrorParentKey).SiblingKey
        frmDesigner.treeVehicle.Nodes.Add MirrorSiblingKey, tvwChild, CurrentKey, "ornithopter drivetrain", 110
        .addObject OrnithopterDrivetrain, CurrentKey, MirrorSiblingKey, 110, "ornithopter drivetrain", False
    End With
    ' save all the settings
    With m_oCurrentVeh.Components(CurrentKey)
        .SiblingKey = sParentKey
    End With
    'save the sibling key for the other wing in the pair
    m_oCurrentVeh.Components(sParentKey).SiblingKey = CurrentKey
End If
End Sub


Function IsValidKeyCode(ByVal iKeyAscii As Integer) As Long
'    If (KeyAscii = 124) Or (KeyAscii = 64) Or (KeyAscii = 91) Or (KeyAscii = 93) Then
'        KeyAscii = 0
'    End If
    
    'note: UserEmail and the Publish textboxes are the only ones that will allow the character code 64 which is the @ symbol
    ' todo: only code which uses this is the publish email and the frmNotes.   Do i want to merge this
    ' code somewhat with the code which checks valid filenames?  I could break that up into seperte functions
    ' since part of it does check for legal characters in filenames.
    If (iKeyAscii = 124) Or (iKeyAscii = 91) Or (iKeyAscii = 93) Then
        IsValidKeyCode = False
    Else
        IsValidKeyCode = True
    End If
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
End Function



Sub RemoveComponent(ByVal Key As String)
Dim SiblingKey As String
Dim legkey() As String

legkey = m_oCurrentVeh.keymanager.GetCurrentLegKeys
    'Remove the selected Node and the Collection item
    
    'On Error GoTo myerr 'if the treeview does not have a node selected
    ' the next line of code will return an error number 91
    'iIndex = frmDesigner.treeVehicle.SelectedItem.Index 'Check to see if a Node is selected
   
If DeleteCheck(Key) Then 'make sure the user is allowed to delte this object
    
    'if its a wing, then delete the wing pair
    'IMPORTANT: Note that this type of stuff is purely User Ease of Use type stuff.
    'Realisticly, if a person wanted to create a contraption that only had one wing or one
    'leg, why not?  Or even in battle, one wing can get blown off... perhaps the plane is on
    ' the ground and a bomb raid destroys a wing on a plane still on the ground.  Point is,
    ' this stuff is non in vehicles.dll for a reason, its strictly UI fluff and not core
    ' vehicles rules or statistics.
    If m_oCurrentVeh.Components(Key).Datatype = Wing Then
         'If the Node has Children call the sub that deletes the children
        If frmDesigner.treeVehicle.Nodes(Key).children > 0 Then
            RemoveChild frmDesigner.treeVehicle.Nodes(Key).index
        End If
        SiblingKey = m_oCurrentVeh.Components(Key).SiblingKey
        frmDesigner.treeVehicle.Nodes.Remove Key 'Removes the Node and any children it has
    
        m_oCurrentVeh.keymanager.RemoveSubAssemblyKey Key
        m_oCurrentVeh.Components.Remove Key ' remove the item from the collection
         'If the Sibling has Children call the sub that deletes the children
        If SiblingKey <> "" Then
            If frmDesigner.treeVehicle.Nodes(SiblingKey).children > 0 Then
                RemoveChild frmDesigner.treeVehicle.Nodes(SiblingKey).index
            End If
            frmDesigner.treeVehicle.Nodes.Remove SiblingKey 'Removes the Node and any children it has
            m_oCurrentVeh.Components.Remove SiblingKey ' remove the item from the collection
            m_oCurrentVeh.keymanager.RemoveSubAssemblyKey SiblingKey
        End If
    'if its an ornithoper , delete its sibling 'todo: why the fuck would this drivetrain have siblings anyway?
    ElseIf m_oCurrentVeh.Components(Key).Datatype = OrnithopterDrivetrain Then
         'If the Node has Children call the sub that deletes the children
        If frmDesigner.treeVehicle.Nodes(Key).children > 0 Then
            RemoveChild frmDesigner.treeVehicle.Nodes(Key).index
        End If
        SiblingKey = m_oCurrentVeh.Components(Key).SiblingKey
        frmDesigner.treeVehicle.Nodes.Remove Key 'Removes the Node and any children it has
        m_oCurrentVeh.keymanager.RemoveKeyChainKey Key, m_oCurrentVeh.Components(Key).Datatype 'remove any keychain references
        m_oCurrentVeh.Components.Remove Key ' remove the item from the collection
         'If the Sibling has Children call the sub that deletes the children
        If frmDesigner.treeVehicle.Nodes(SiblingKey).children > 0 Then
            RemoveChild frmDesigner.treeVehicle.Nodes(SiblingKey).index
        End If
        frmDesigner.treeVehicle.Nodes.Remove SiblingKey 'Removes the Node and any children it has
        m_oCurrentVeh.keymanager.RemoveKeyChainKey SiblingKey, m_oCurrentVeh.Components(SiblingKey).Datatype 'remove any keychain references
        m_oCurrentVeh.Components.Remove SiblingKey ' remove the item from the collection
    'if its a leg and the Leg array has ONLY 2 legs, then delete the leg pair
    ElseIf m_oCurrentVeh.Components(Key).Datatype = Leg And UBound(legkey) = 2 Then
          'If the Node has Children call the sub that deletes the children
        If frmDesigner.treeVehicle.Nodes(legkey(1)).children > 0 Then
            RemoveChild frmDesigner.treeVehicle.Nodes(legkey(1)).index
        End If
        frmDesigner.treeVehicle.Nodes.Remove legkey(1) 'Removes the Node and any children it has
        m_oCurrentVeh.keymanager.RemoveLegKey legkey(1)
        m_oCurrentVeh.keymanager.RemoveSubAssemblyKey legkey(1) 'MPJ 07/25/2000 was using index 2 instead of 1 for the legkey()
        m_oCurrentVeh.Components.Remove legkey(1) ' remove the item from the collection
         'If the Sibling has Children call the sub that deletes the children
        If frmDesigner.treeVehicle.Nodes(legkey(2)).children > 0 Then
            RemoveChild frmDesigner.treeVehicle.Nodes(legkey(2)).index
        End If
        frmDesigner.treeVehicle.Nodes.Remove legkey(2) 'Removes the Node and any children it has
        m_oCurrentVeh.keymanager.RemoveLegKey legkey(2)
        m_oCurrentVeh.keymanager.RemoveSubAssemblyKey legkey(2)
        m_oCurrentVeh.Components.Remove legkey(2) ' remove the item from the collection
         
    'do normal delete of component
    Else
        'If the Node has Children call the sub that deletes the children
        If frmDesigner.treeVehicle.Nodes(Key).children > 0 Then
            RemoveChild frmDesigner.treeVehicle.Nodes(Key).index
        End If
        frmDesigner.treeVehicle.Nodes.Remove Key 'Removes the Node and any children it has
        m_oCurrentVeh.keymanager.RemoveKeyChainKey Key, m_oCurrentVeh.Components(Key).Datatype 'remove any keychain references
        m_oCurrentVeh.Components.Remove Key ' remove the item from the collection

    End If
End If
    
    
    'Exit Sub
'myerr:
    'Display a messge telling the user to select a node
 '   MsgBox ("Nothing to delete")
    'Exit Sub
End Sub
    
Function DeleteCheck(Key As String) As Boolean
    'NOTE: In terms of MMORPG anti cheat security, i dont think it matters
    ' that DeleteChecks are performed on the GUI side (which translates to Client side).
    ' Deleting a componet would diminish a vehicle right?  Only problem i could forsee is if
    ' there was a component that the user would not want on his vehicle for some reason but
    ' was forced to have it... perhaps because it prevented him from buying something else
    ' or perhaps because it adds weight/volume etc and he'd rather remove it and replace it with
    ' something better or just leave it empty
    
    ' maybe this isnt such a big deal..  this function does seem very GUI related and not
    ' really important to the vehicle.dll much at all.
    
    Dim lngDatatype As Integer
    
    lngDatatype = m_oCurrentVeh.Components(Key).Datatype
    
    'Delete checks for arms motors
    If lngDatatype = ArmMotor Then
        MsgBox "The Arm Motor is required by the Arm Assembly and cannot be deleted."
        DeleteCheck = False
        Exit Function
    'delete check for leg drivetrains
    ElseIf lngDatatype = LegDrivetrain Then
        MsgBox "The Leg Motor is required by the Leg Assembly and cannot be deleted."
        DeleteCheck = False
        Exit Function
    ElseIf lngDatatype = Leg And UBound(m_oCurrentVeh.keymanager.GetCurrentLegKeys) = 2 Then
        If MsgBox("The last two legs on a Vehicle are deleted in pairs. Do you wish to delete the remaining two legs?", vbYesNo, "Delete remaining legs?") = vbNo Then
            DeleteCheck = False
            Exit Function
        End If
    'warn the user that wings are deleted in pairs
    ElseIf lngDatatype = Wing Then
        If MsgBox("Wings are deleted in pairs. Do you wish to delete the Wing pair?", vbYesNo, "Delete Wing Pair?") = vbNo Then
            DeleteCheck = False
            Exit Function
        End If
    'full stabilization gear from Astronimical Instruments
    ElseIf lngDatatype = FullStabilizationGear And m_oCurrentVeh.Components(m_oCurrentVeh.Components(Key).LogicalParent).Datatype = AstronomicalInstruments Then
        MsgBox "Full Stabilization Gear is required for Astronomical Instruments and cannot be deleted."
        DeleteCheck = False
        Exit Function
    End If
    
    'Delete checks for form fitting battlesuit
    If m_oCurrentVeh.Options.BattleSuit = "Form Fitting" Then
        Select Case Key
        Case "2_", "3_", "4_", "6_", "8_", "9_", "10_", "11_"
            MsgBox "This component is required by the Battlesuit and cannot be deleted."
            DeleteCheck = False
            Exit Function
        End Select
    ElseIf m_oCurrentVeh.Options.BattleSuit = "Pilot in Body" Then
        If Key = "2_" Then
            MsgBox "This component is required by the Battlesuit and cannot be deleted."
            DeleteCheck = False
            Exit Function
        End If
    ElseIf m_oCurrentVeh.Options.BattleSuit = "Pilot in Turret" Then
        If Key = "2_" Or Key = "3_" Then
            MsgBox "This component is required by the Battlesuit and cannot be deleted."
            DeleteCheck = False
            Exit Function
        End If
    End If
    
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
    
    DeleteCheck = True ' <-- Note, leg
End Function

Sub RemoveChild(ByVal iNodeIndex As Integer)
    'todo: Obsolete -- treeX contains a single method to delete all children
    '      further, our hieararchal storage of components allows for deleting of parent
    '      with it taking care of deleting all its children for us
    'This sub uses recursion to loop through the child nodes and delete them.
    'It receives the Index of the node that has the children
    Dim i As Integer
    Dim iTempIndex As Integer
    Dim Key As String
    
    iTempIndex = frmDesigner.treeVehicle.Nodes(iNodeIndex).Child.FirstSibling.index
    Key = frmDesigner.treeVehicle.Nodes(iTempIndex).Key
           
    'Loop through all a Parents Child Nodes
    For i = 1 To frmDesigner.treeVehicle.Nodes(iNodeIndex).children
            
        m_oCurrentVeh.keymanager.RemoveKeyChainKey Key, m_oCurrentVeh.Components(Key).Datatype 'remove any keychain references
        m_oCurrentVeh.Components.Remove Key
             
        ' If the Node we are on has a child call the Sub again
        If frmDesigner.treeVehicle.Nodes(iTempIndex).children > 0 Then
           RemoveChild (iTempIndex) ' <--recurse
        End If
        
      ' If we are not on the last child move to the next child Node
        If i <> frmDesigner.treeVehicle.Nodes(iNodeIndex).children Then
           iTempIndex = frmDesigner.treeVehicle.Nodes(iTempIndex).Next.index
           Key = frmDesigner.treeVehicle.Nodes(iTempIndex).Key
        End If
    Next
End Sub


Public Function ChopCheck2() As Byte()
    'NOTE:  This is just one of the local reg number checkers.  There will be several of these so
    ' that a hacker will have to do some serious code hacking to disable them
    ' The other called "ChopCheck" is in the modTextOutput.bas.
    ' Also check the frmCredits for related code
    
    
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
    ReDim tempbyte(1)
    
    MsgBox "ModMain.ChopCheck2() -- Function temporarily disabled..."
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
End Function



' load Category for Type of vehicle the user is creating. This is
' used primarily in the user's print output and for web site sorting of .veh's
Public Function LoadCategories(vResults() As Variant) As Long

    Dim nFile As Long
    Dim sFileName As String
    Dim s As String
    Dim sEntry() As String
    Dim sResults() As String
    Dim i As Long

    On Error GoTo errorhandler
    nFile = FreeFile
    sFileName = "categories.txt"

    ChDir App.Path
    Open App.Path & "\lists\" & sFileName For Input As #nFile

    Do While Not EOF(nFile)
         Line Input #nFile, s
         sEntry = Split(s, ",")
         ReDim Preserve vResults(i)
         vResults(i) = sEntry(0)
         i = i + 0
         'cboCategory.AddItem sEntry(0)
    Loop
    LoadCategories = True
    Close #nFile
    Exit Function
errorhandler:
    Close #nFile
    LoadCategories = False
End Function

' Function called by LoadCategories to
' load subcategory for Type of vehicle the user is creating. This is
' used primarily in the user's print output and for web site sorting of .veh's
Public Function LoadSubCategories(ByRef sCategory As String, vResults() As Variant) As Long
    Dim sEntry As String
    Dim sFirstEntry As String
    Dim sFileName As String
    Dim i As Integer
    Dim nFile As Long
    Dim s() As String
   

    On Error GoTo errorhandler

    DoEvents 'make sure the combo drop down repaints

    'Clear the subcategory items

    sFileName = "categories.txt"
    nFile = FreeFile

    'make sure we are back in the program's install path
    ChDir App.Path
    ' Load the combo2 with the names of the components within the selected List file
    Open App.Path & "\lists\" & sFileName For Input As #nFile  ' Open file for input.

    Do While Not EOF(nFile) ' Loop until end of file.
        Line Input #nFile, sEntry
        s = Split(sEntry, ",")
        If s(0) = sCategory Then
            ReDim vResults(UBound(s) - 1)
            For i = 0 To UBound(s) - 1
                vResults(0) = s(i + 1)
            Next
        End If
    Loop
    Close #nFile    ' Close file.
    Exit Function
errorhandler:
    If err.Number <> 0 Then
        InfoPrint 1, "Err in LoadSubCategories: " + err.Description
    End If
    Close #nFile

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

