VERSION 5.00
Begin VB.Form frmDesignCheck 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Design Check"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6300
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDesignCheck 
      Height          =   4545
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   30
      Width           =   6240
   End
End
Attribute VB_Name = "frmDesignCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Dim sRet As String
    Dim sTotal As String
    
    'TODO: this was removed from NewProfile and needs to be added somewhere
    'if user chose to make Submerged, check to make sure Vehicle Options Submerged is true
    'If cboPerformanceType.Text = "Submerged" And m_oCurrentVeh.Components(BODY_KEY).Submersible <> True Then
    '    MsgBox "You have not yet enabled the Submersible property in the Options dialog.  This will be done for you."
    '    m_oCurrentVeh.Components(BODY_KEY).Submersible = True
    'End If
    
    
    sTotal = ""
    sTotal = MinBodyVolumeCheck
    MsgBox "Design Check not yet implemented..."
    Exit Sub
    ' check that weights loaded onto hardpoitns dont exceed the hardpoints' capacity
    sRet = HardpointCapacityCheck
    Call AppendCheckText(sTotal, sRet)
    
    ' check leg volumes match and that they meet minimum volume reqts
    sRet = LegCheck
    Call AppendCheckText(sTotal, sRet)
    
    
    sRet = SoftwareComplexityCheck
    Call AppendCheckText(sTotal, sRet)
    
    ' check that if a duplicate maneuver control is found, that an original (non duplicate) one exists as well
    sRet = DuplicateControlsCheck
    Call AppendCheckText(sTotal, sRet)
    
    ' check that all subs with have armor if streamlining is used
    sRet = ArmorCheck
    Call AppendCheckText(sTotal, sRet)
    
    ' check max non rigid armor DR's are not exceeded
    sRet = MaxNonRigidArmorDRCheck
    Call AppendCheckText(sTotal, sRet)
    
    ' check that engines are connected to fuel tanks which use the same type of fuel
    sRet = EngineTankFuelTypeCheck
    Call AppendCheckText(sTotal, sRet)
    
    ' check for hardpoints on body exception rules
    sRet = HardpointsOnBodyCheck
    Call AppendCheckText(sTotal, sRet)
    
    ' check for streamlining rules violations
    sRet = StreamliningViolationsCheck
    Call AppendCheckText(sTotal, sRet)
    
    ' check for recommended minimum wing volumes
    sRet = WingVolumesCheck
    Call AppendCheckText(sTotal, sRet)
    
    ' check for solar cell placement violations
    sRet = SolarCellArrayCheck
    Call AppendCheckText(sTotal, sRet)
    
    ' inform user for reasons of any speed limits they may have hit
    sRet = TopSpeedRestrictions
    Call AppendCheckText(sTotal, sRet)
    
    sRet = StackedTurretsCheck
    Call AppendCheckText(sTotal, sRet)
    
    ' Design Check:  Make sure vehicles Max Lift or Floatation rating has not been exceeded
' TODO If it does tell user to convert cargo space to empty space
' to reduce its weight or eliminate some armor.
    txtDesignCheck = sTotal
    
    If txtDesignCheck = "" Then txtDesignCheck = "No design flaws detected in current vehicle."
    
End Sub

Sub AppendCheckText(ByRef sTarget As String, ByRef sChunk As String)

    If sChunk <> "" Then
    
        If sTarget = "" Then
            sTarget = sChunk
        Else
            sTarget = sTarget & vbNewLine & sChunk
        End If
    End If
    

End Sub

Function StackedTurretsCheck() As String
    Dim sParent As String
    Dim Temp As String
    Dim subsarray() As String
    Dim i, num As Long
    Dim sRet As String
     
    On Error Resume Next
     
    ' get all our subassemblies.  If we have none except the body, exit function
    subsarray = m_oCurrentVeh.Components(BODY_KEY).GetCurrentSubAssembliesKeys
    num = UBound(subsarray)
    If subsarray(1) = "" Then Exit Function
    
    ' MPJ 11/7/2000
    ' this has been removed as a hard restriction in the
    ' turrets Let Orientation() property and is
    ' place here in the design check instead
    For i = 1 To num
        If TypeOf m_oCurrentVeh.Components(subsarray(i)) Is clsTurret Then
            sParent = m_oCurrentVeh.Components(subsarray(i)).Parent  'note: NOT logical parent since subassemblies cant be attached to Groups
            If TypeOf m_oCurrentVeh.Components(sParent) Is clsTurret Then
            
                Temp = m_oCurrentVeh.Components(sParent).Orientation
                If Temp <> m_oCurrentVeh.Components(subsarray(i)).Orientation Then
                    sRet = "FYI: Stacked Turrets should usually share the same orientation."
                End If
            End If
        End If
    Next
    
    StackedTurretsCheck = sRet
End Function

Function WingVolumesCheck() As String
    'checks that wings are at the recommended 0.1 x body volume and for stub wings at least 0.02 x body volume
    
    Dim element As Object
    Dim dblBodyVolume As Double
    Dim dblWingVolume As Double
    Dim sRet As String
    Dim dblMinStub As Double
    Dim dblMinStandard As Double
    
    dblBodyVolume = m_oCurrentVeh.Components(BODY_KEY).Volume
    
    dblMinStandard = Round(0.1 * dblBodyVolume, 2)
    dblMinStub = Round(0.02 * dblBodyVolume, 2)
    
    For Each element In m_oCurrentVeh.Components
        If TypeOf element Is clsWing Then
        
            dblWingVolume = element.Volume
            
            If element.subtype = "stub" Then
                If dblWingVolume < dblMinStub Then
                    sRet = "WARNING: Detected a wing volume of " & dblWingVolume & ".  Minimum recommended volume for 'stub wings' is 0.02 x Body Volume. In your case, each stub wing should be at least " & dblMinStub & " cf."
                End If
            Else
                If dblWingVolume < dblMinStandard Then
                    sRet = "WARNING: Detected a wing volume of " & dblWingVolume & ".  Minimum recommended volume for wings is 0.1 x Body Volume. In your case, each wing should be at least " & dblMinStandard & " cf."
                End If
            End If
        End If
    Next
    
    WingVolumesCheck = sRet
End Function


Function StreamliningViolationsCheck() As String
    'VE page 11, "a vehicle with masts cannot have better than Fair streamlining.
    ' A vehicle with skids or wheels (unless retractable), tracks, halftracks, skitraacks
    ' biplane or triaplane wings, GEV skirs, SEV sidewalls, open mounts, gasbags, rotors,
    ' arms or legs cannot have better than Good streamling.  A vehicle with superstructures
    ' or turrets (except pop turrets) cannot have better than Very Good streamling. Vehicles
    ' with wings cannot have Superior, Excellent or Radical streamlining before TL7.
    
    
    Dim sType As String
    Dim lngDType As Long
    Dim subsarray() As String
    Dim i, num As Long
    Dim sRet As String
     
    On Error Resume Next
     
    ' get all our subassemblies.  If we have none except the body, exit function
    subsarray = m_oCurrentVeh.Components(BODY_KEY).GetCurrentSubAssembliesKeys
    num = UBound(subsarray)
    If subsarray(1) = "" Then Exit Function
                
    ' determine which level of streamlining the user has
    sType = m_oCurrentVeh.surface.StreamLining
    
    Select Case sType
        Case "none", "fair"
            Exit Function
            
        Case "good", "very good", "superior", "excellent", "radical"
            ' check for masts limit of "fair" streamlining
            For i = 1 To num
                lngDType = m_oCurrentVeh.Components(subsarray(i)).Datatype
            
                Select Case lngDType
                    Case Mast
                        sRet = "WARNING: You have Masts installed and " & sType & " streamlining. Vehicles with Masts cant have better than 'fair' streamlining."

                End Select
            Next
            
            ' check for rules which limit streamlining to "good" or lower
            If sType <> "good" Then
                For i = 1 To num
                    lngDType = m_oCurrentVeh.Components(subsarray(i)).Datatype
                
                    Select Case lngDType
                        
                        Case Skid ' unless retractable
                            If (m_oCurrentVeh.Components(subsarray(i)).RetractLocation = "none") Then
                                If sRet <> "" Then sRet = sRet & vbNewLine
                                sRet = sRet & "WARNING: Your vehicle has non-retractable skids installed and " & sType & " streamlining.  Vehicles with with non-retractable skids cannot have better than 'good' streamlining."
                            End If
                            
                        Case Wheel 'unless retractable
                            If (m_oCurrentVeh.Components(subsarray(i)).subtype <> "retractable") Then
                                If sRet <> "" Then sRet = sRet & vbNewLine
                                sRet = sRet & "WARNING: Your vehicle has non-retractable wheels installed and " & sType & " streamlining.  Vehicles with with non-retractable wheels cannot have better than 'good' streamlining."
                            End If
                            
                        Case Track
                            If sRet <> "" Then sRet = sRet & vbNewLine
                            sRet = sRet & "WARNING: Your vehicle has a Track subassembly installed and " & sType & " streamlining.  Vehicles with with Tracks cannot have better than 'good' streamlining."
                            
                            
                        Case Wing ' biplane or triaplane wings only
                            If (m_oCurrentVeh.Components(subsarray(i)).subtype = "biplane") Or (m_oCurrentVeh.Components(subsarray(i)).subtype = "triplane") Then
                                If sRet <> "" Then sRet = sRet & vbNewLine
                                sRet = sRet & "WARNING: Your vehicle has biplane or triplane wings installed and " & sType & " streamlining.  Vehicles with bi/triplane wings cannot have better than 'good' streamlining."
                            End If
                        Case Hovercraft
                            If sRet <> "" Then sRet = sRet & vbNewLine
                            sRet = sRet & "WARNING: Your vehicle has a GEV or SEV hovercraft subassembly installed and " & sType & " streamlining.  Vehicles with with GEV or SEV cannot have better than 'good' streamlining."
                            
                        Case OpenMount
                            If sRet <> "" Then sRet = sRet & vbNewLine
                            sRet = sRet & "WARNING: Your vehicle has Open Mount subassemblies installed and " & sType & " streamlining.  Vehicles with with Open Mounts cannot have better than 'good' streamlining."
                            
                        Case Gasbag
                            If sRet <> "" Then sRet = sRet & vbNewLine
                            sRet = sRet & "WARNING: Your vehicle has a Gasbag subassembly installed and " & sType & " streamlining.  Vehicles with with Gasbags cannot have better than 'good' streamlining."
        
                        Case AutogyroRotor, TTRotor, CARotor, MMRotor
                            If sRet <> "" Then sRet = sRet & vbNewLine
                            sRet = sRet & "WARNING: Your vehicle has a Rotor subassembly installed and " & sType & " streamlining.  Vehicles with with Rotors cannot have better than 'good' streamlining."
                            
                        Case Arm
                             If sRet <> "" Then sRet = sRet & vbNewLine
                            sRet = sRet & "WARNING: Your vehicle has an Arm subassembly installed and " & sType & " streamlining.  Vehicles with with Arms cannot have better than 'good' streamlining."
                            
                        Case Leg
                             If sRet <> "" Then sRet = sRet & vbNewLine
                            sRet = sRet & "WARNING: Your vehicle has a Leg subassembly installed and " & sType & " streamlining.  Vehicles with with Legs cannot have better than 'good' streamlining."
                            
                       
                    End Select
                Next
                
            ' check for rules which limit streamlining to "very good" or lower
            ElseIf sType <> "very good" Then
                For i = 1 To num
                    lngDType = m_oCurrentVeh.Components(subsarray(i)).Datatype
                
                    Select Case lngDType
                        Case Turret
                            If sRet <> "" Then sRet = sRet & vbNewLine
                            sRet = sRet & "WARNING: You have turrets installed and " & sType & " streamlining.  Vehicles with turrets cannot have better than "
                       
                        Case Superstructure
                            If sRet <> "" Then sRet = sRet & vbNewLine
                            sRet = sRet & "WARNING: You have superstructures installed and " & sType & " streamlining.  Vehicle with superstructures cannot have better than 'very good' streamlining."
                        
                        Case Wing
                            If m_oCurrentVeh.Components(BODY_KEY).TL < 7 Then
                            ' if it has wings and before TL 7, then it cannot have superior or better streamlining
                  
                                If sRet <> "" Then sRet = sRet & vbNewLine
                                sRet = sRet & "WARNING: You have wings installed and " & sType & " streamlining.  Vehicles prior to TL7 cannot have better than 'very good' streamlining."
                            End If
                    End Select
                Next
            End If
            
    End Select
    
    StreamliningViolationsCheck = sRet

End Function

Function HardpointsOnBodyCheck() As String
    'page 94 - hardpoints may not be added to a vehicle's body if it has a hydrodynamic hull,
    'GEV or SEV subassemblies, tracks, halftracks or skitracks, railway wheels or a
    'flexibody drivetrain.
    
    Dim element As Object
    Dim bHydro As Boolean
    Dim bGEVSEV As Boolean
    Dim bTracks As Boolean
    Dim bRailwayWheels As Boolean
    Dim bFlexibody As Boolean
    Dim bHardpointsOnBody As Boolean
    Dim sRet As String
    
    On Error Resume Next
    
    ' determine if we have hardpoints on the body
    For Each element In m_oCurrentVeh.Components
        If TypeOf element Is clsHardPoint Then
            If element.Datatype = HardPoint Then
                bHardpointsOnBody = True
                Exit For
            End If
        End If
    Next
    
    ' if we do have hardpoints on body, check to see if exceptions are found which
    ' recommend hardpoints not be added to body
    If bHardpointsOnBody Then
        For Each element In m_oCurrentVeh.Components
            If TypeOf element Is clsBody Then
                If element.HydrodynamicLines <> "none" Then
                    bHydro = True
                End If
            ElseIf TypeOf element Is clsHovercraft Then
                bGEVSEV = True
            ElseIf TypeOf element Is clsTrack Then
                bTracks = True
            ElseIf TypeOf element Is clsWheel Then
                If element.subtype = "railway" Then
                    bRailwayWheels = True
                End If
            ElseIf TypeOf element Is clsGroundDrivetrain Then
                If element.Datatype = FlexibodyDrivetrain Then
                    bFlexibody = True
                End If
            End If
        Next
    Else
        Exit Function
    End If
    
    ' create print output
    If bHydro Then sRet = sRet & "hydrodynamic lines, "
    If bGEVSEV Then sRet = sRet & "a GEV or SEV hovercraft subassembly, "
    If bTracks Then sRet = sRet & "tracks, "
    If bRailwayWheels Then sRet = sRet & "railway wheels, "
    If bFlexibody Then sRet = sRet & "a flexibody drivetrain "
    
    If sRet <> "" Then
        sRet = "WARNING: You have hardpoints attached to the body. You also have " & sRet
        sRet = sRet & ". VE 2nd edition suggests that hardpoints not be added to the body IF it has hydrodynamic lines " _
            & "GEV or SEV subassemblies, tracks, halftracks, or skitracks, railways wheels or a flexibody drivetrain."
    End If
    
    HardpointsOnBodyCheck = sRet
End Function

Function EngineTankFuelTypeCheck() As String

    Dim arrEngines() As String
    Dim arrTanks() As String
    Dim i As Long
    Dim j As Long
    Dim sText As String
    Dim sFuelType As String
    Dim sTankDescription As String
    Dim sEngineDescription As String
    
    On Error Resume Next
    
    arrTanks = m_oCurrentVeh.Components(BODY_KEY).GetCurrentFuelStorageKeys
    If arrTanks(1) = "" Then Exit Function
    
    For i = 1 To UBound(arrTanks)
        sFuelType = m_oCurrentVeh.Components(arrTanks(i)).Fuel
        sTankDescription = m_oCurrentVeh.Components(arrTanks(i)).CustomDescription
        
        arrEngines = m_oCurrentVeh.Components(BODY_KEY).GetCurrentFuelUsingSystemKeys
        If arrEngines(1) <> "" Then
            For j = 1 To UBound(arrEngines)
                sEngineDescription = m_oCurrentVeh.Components(arrEngines(j)).CustomDescription
                If m_oCurrentVeh.Components(arrEngines(j)).Fueltype <> sFuelType Then
                    sText = sText & "Fuel type for " & sEngineDescription & " does not match fuel stored in " & sTankDescription & ".  "
                End If
            Next
        End If
    Next
    
    
    EngineTankFuelTypeCheck = sText

End Function

Function MaxNonRigidArmorDRCheck() As String

    Dim element As Object
    Dim dType As Long
    Dim sText As String
    Dim bReflex As Boolean
    Dim bFlag As Boolean
    
    For Each element In m_oCurrentVeh.Components
        If TypeOf element Is clsArmor Then
            dType = element.Datatype
            
            Select Case dType
            
                Case ArmorBasicFacing, ArmorOverall, ArmorWheelGuard, ArmorGunShield, ArmorOpenFrame, ArmorLocation, ArmorComponent
                    If element.material = "nonrigid" Then
                        bReflex = True
                        If element.dr > 100 Then
                            bFlag = True ' bail out as soon as we violate the rule
                            Exit For
                        End If
                    End If
                    
                Case ArmorComplexFacing
                    If element.material1 = "nonrigid" Then
                         bReflex = True
                        If element.dr1 > 100 Then
                            bFlag = True
                            Exit For
                        End If
                    ElseIf element.material2 = "nonrigid" Then
                         bReflex = True
                        If element.dr2 > 100 Then
                            bFlag = True
                            Exit For
                        End If
                    ElseIf element.material3 = "nonrigid" Then
                         bReflex = True
                        If element.dr3 > 100 Then
                            bFlag = True
                            Exit For
                        End If
                    ElseIf element.material4 = "nonrigid" Then
                         bReflex = True
                        If element.dr4 > 100 Then
                            bFlag = True
                            Exit For
                        End If
                    ElseIf element.material5 = "nonrigid" Then
                         bReflex = True
                        If element.dr5 > 100 Then
                            bFlag = True
                            Exit For
                        End If
                    ElseIf element.material6 = "nonrigid" Then
                         bReflex = True
                        If element.dr6 > 100 Then
                            bFlag = True
                            Exit For
                        End If
                    End If
                    
            End Select
        End If
    Next
        
    If bFlag Then
    
        sText = "WARNING: You have set DR greater than 100 for an armor component which uses 'nonrigid' material.  Some GURPS users feel nonrigid armor DR cannot exceed 100."
    End If
    
    If bReflex Then
        If sText <> "" Then
            sText = sText & vbNewLine & "NOTE: If you intend for your nonrigid armor to be 'reflex' armor, the DR should be limited to 5 x TL (VE page 22)"
        Else
        
        
            sText = "NOTE: If you intend for your nonrigid armor to be 'reflex' armor, the DR should be limited to 5 x TL (VE page 22)"
        End If
    End If

    MaxNonRigidArmorDRCheck = sText
    
End Function

Function ArmorCheck() As String
    
    Dim bSubmersible As Boolean
    Dim bFloatationHull As Boolean
    Dim bStreamlining As Boolean
    Dim bSlope As Boolean
    Dim bRotors As Boolean
    Dim arrSubs() As String
    Dim j As Long
    Dim sKey As String
    Dim bCurrentSubOK As Boolean
    Dim sText As String
 
    
    Dim element As Object
    
    With m_oCurrentVeh.surface
        bSubmersible = .Submersible
        bFloatationHull = .FloatationHull
        If .StreamLining <> "none" Then
            bStreamlining = True
        End If
    End With
    
    With m_oCurrentVeh.Components(BODY_KEY)
        ' check for slope on body
        If .SlopeR <> "none" Then
            bSlope = True
        ElseIf .slopel <> "none" Then
            bSlope = True
        ElseIf .slopef <> "none" Then
            bSlope = True
        ElseIf .slopeb <> "none" Then
            bSlope = True
        End If
    End With
    
    ' get array of all subassembly keys
    arrSubs = m_oCurrentVeh.keymanager.GetCurrentSubAssembliesKeys
    
        
' check if any of these subs is a rotor or has slope
If arrSubs(1) <> "" Then
    For j = 1 To UBound(arrSubs)
        If TypeOf m_oCurrentVeh.Components(arrSubs(j)) Is clsrotor Then
            bRotors = True
        Else
            ' check to see if the vehilce has slope anywhere
            If Not bSlope Then
                If (TypeOf m_oCurrentVeh.Components(arrSubs(j)) Is clsBody) Or _
                    (TypeOf m_oCurrentVeh.Components(arrSubs(j)) Is clsSuperStructure) Or _
                    (TypeOf m_oCurrentVeh.Components(arrSubs(j)) Is clsTurret) Or _
                    (TypeOf m_oCurrentVeh.Components(arrSubs(j)) Is clsPopTurret) Then
                    
                    With m_oCurrentVeh.Components(arrSubs(j))
                        If .SlopeR <> "none" Then
                            bSlope = True
                        ElseIf .slopel <> "none" Then
                            bSlope = True
                        ElseIf .slopef <> "none" Then
                            bSlope = True
                        ElseIf .slopeb <> "none" Then
                            bSlope = True
                        End If
                    End With
                End If
            End If
            
        ' bail out of For Loop early if possible
        If bSlope And bRotors Then Exit For
        End If
    Next
End If

'now loop thru all subassemblies and make sure they have armor
If arrSubs(1) <> "" Then
    For j = 1 To UBound(arrSubs)
        sKey = m_oCurrentVeh.Components(arrSubs(j)).Key
        
        For Each element In m_oCurrentVeh.Components
        
            If TypeOf element Is clsArmor Then
                If element.Datatype = ArmorOverall Then
                    Exit Function
                
                ElseIf element.LogicalParent = sKey Then
                    bCurrentSubOK = True ' this subassembly does have armor. Keep checking though to see if we find an "overall armor"
                
            
                End If
            End If
        Next
        
        If Not bCurrentSubOK Then
            Exit For ' bail out.  We do not have armor on all subs
        Else
            bCurrentSubOK = False ' reset and test next sub
        End If
    Next
End If
    
' if we havented exited this function, then we do not have necessary armor
' so we generate our message
If bRotors Then sText = " rotors,"
If bSubmersible Then sText = sText & " submersible hull,"
If bFloatationHull Then sText = sText & " floatation hull," 'todo: floationhull should be determined if the floatationrating is greater than 0!  This depends
If bStreamlining Then sText = sText & " streamlining,"
If bSlope Then sText = sText & " slope,"

If sText = "" Then Exit Function

sText = "Your vehicle has " & sText & " this requires all subassemblies to be armored or that 'overall armor' is used. Please add armor."
    
ArmorCheck = sText
    
End Function


Function DuplicateControlsCheck() As String
    
    Dim bMechanical As Boolean
    Dim bMechanicalDup As Boolean
    Dim bMechanicalDiving As Boolean
    Dim bMechanicalDivingDup As Boolean
    
    
    Dim bComputerized As Boolean
    Dim bComputerizedDup As Boolean
    Dim bComputerizedDiving As Boolean
    Dim bComputerizedDivingDup As Boolean
    
    
    Dim bElectric As Boolean
    Dim bElectricDup As Boolean
    Dim bElectricDiving As Boolean
    Dim bElectricDivingDup As Boolean
    
    Dim element As Object
    Dim sRet As String
    
    On Error Resume Next
    
    For Each element In m_oCurrentVeh.Components
        If TypeOf element Is clsManeuverControl Then
            Select Case element.Datatype
                
                Case ElectronicDivingControl
                    If element.duplicate Then
                    
                        bElectricDivingDup = True
                    Else
                        bElectricDiving = True
                    End If
                    
                Case ElectronicManeuverControl
                    If element.duplicate Then
                        bElectricDup = True
                    Else
                        bElectric = True
                    End If
                    
                Case ComputerizedDivingControl
                    If element.duplicate Then
                        bComputerizedDivingDup = True
                    Else
                        bComputerizedDiving = True
                    End If
                    
                Case ComputerizedManeuverControl
                    If element.duplicate Then
                        bComputerizedDup = True
                    Else
                        bComputerized = True
                    End If
                    
                Case MechanicalManeuverControl
                    If element.duplicate Then
                        bMechanicalDup = True
                    Else
                        bMechanical = True
                    End If
                    
                Case MechanicalDivingControl
                    If element.duplicate Then
                        bMechanicalDivingDup = True
                    Else
                        bMechanicalDiving = True
                    End If
            End Select
        End If
    Next
    
    
    If (bMechanicalDivingDup) And (bMechanicalDiving = False) Then
        sRet = "You have a duplicate Mechanical Diving control but no primary (non duplicate) set installed."
    End If
    
    If (bMechanicalDup) And (bMechanical = False) Then
        If sRet <> "" Then sRet = sRet & vbNewLine
        sRet = sRet & "You have a duplicate Mechanical maneuver control but no primary (non duplicate) set installed."
    End If
    
    If (bComputerizedDup) And (bComputerized = False) Then
        If sRet <> "" Then sRet = sRet & vbNewLine
        sRet = sRet & "You have a duplicate Computerized maneuver control but no primary (non duplicate) set installed."
    End If
    
    If (bComputerizedDivingDup) And (bComputerizedDiving = False) Then
        If sRet <> "" Then sRet = sRet & vbNewLine
        sRet = sRet & "You have a duplicate Computerized diving control but no primary (non duplicate) set installed."
    End If
    
    If (bElectricDup) And (bElectric = False) Then
        If sRet <> "" Then sRet = sRet & vbNewLine
        sRet = sRet & "You have a duplicate Electronic maneuver control but no primary (non duplicate) set installed."
    End If
    
    If (bElectricDivingDup) And (bElectricDiving = False) Then
        If sRet <> "" Then sRet = sRet & vbNewLine
        sRet = sRet & "You have a duplicate Electronic diving control but no primary (non duplicate) set installed."
    End If
    
   DuplicateControlsCheck = sRet
End Function

Function MinBodyVolumeCheck() As String
'    Dim TempVolume As Single
'    Dim subsarray() As String
'    Dim tempvolume2 As Single
'    Dim i, num As Long
'    Dim minvolume As Single
'    Dim retval As String
'     On Error Resume Next
'
'    subsarray = m_oCurrentVeh.Components(BODY_KEY).GetCurrentSubAssembliesKeys
'    num = UBound(subsarray)
'    If subsarray(1) <> "" Then
'        For i = 1 To num
'            Select Case m_oCurrentVeh.Components(subsarray(i)).Datatype
'                'determine our minimum volume by adding combined volume of all
'                'turrets, arms, open mounts, superstructers, and pods directly
'                'atached to the body.
'                Case Turret, Popturret, Arm, OpenMount, Superstructure, Pod
'                    TempVolume = TempVolume + m_oCurrentVeh.Components(subsarray(i)).Volume
'                ' Determine minimum volume due to masts
'                Case Mast
'                    tempvolume2 = (m_oCurrentVeh.Components(subsarray(i)).Height / 4) ^ 3
'            End Select
'        Next
'    End If
'
'    '//store the minimum volume
'    minvolume = Maximum(tempvolume2, TempVolume)
'    'check the real body volume must exceed the minimum. If it doesnt, add empty space.
'    If TempVolume >= tempvolume2 Then
'        If (TempVolume > 0) And (TempVolume > m_oCurrentVeh.Components(BODY_KEY).Volume) Then
'            retval = "Real body volume must exceed combined volume of all turrets, arms, " _
'            & "open mounts, superstructures, and pods directly attached to the body.  Add empty space."
'        End If
'    ' Determine if the minimum volume due to masts is achieved
'    ElseIf (tempvolume2 > 0) And (tempvolume2 > m_oCurrentVeh.Components(BODY_KEY).Volume) Then
'        retval = "Real body volume is currently " & Round(m_oCurrentVeh.Components(BODY_KEY).Volume, 2) & " It must be greater or equal to " & Round(minvolume, 2) & " [Tallest Mast Height /4) cubed].  You must add at least " & Round(minvolume - m_oCurrentVeh.Components(BODY_KEY).Volume, 2) & " of empty space."
'    End If
'
'    MinBodyVolumeCheck = retval
End Function

Function LegCheck() As String
    Dim MinLegVolume As Single ' each leg must be (0.04 x body volume / number of legs) in volume
    Dim arrLegVolumes() As Single ' array to hold the volumes for all the legs
    Dim legarray() As String 'holds the keys for all legs in the vehicle
    Dim i, NumLegs, j As Integer ' counter
    Dim retval As String
    Dim retval2 As String
    
    On Error GoTo errorhandler
    ' Find how many legs are on the vehicle
    legarray = m_oCurrentVeh.Components(BODY_KEY).GetCurrentLegKeys
    NumLegs = UBound(legarray)
    If legarray(1) <> "" Then
        ReDim arrLegVolumes(NumLegs) 'set the dimension of the array to the number of legs
        ' Check to see if Leg minimum volume is attained
        MinLegVolume = (m_oCurrentVeh.Components(BODY_KEY).Volume * 0.04 / NumLegs)
        MinLegVolume = Round(MinLegVolume, 2)
        
        ' set the volume of the first leg which will be used to compare volumes of others
        arrLegVolumes(1) = m_oCurrentVeh.Components(legarray(1)).Volume
        
        'check to see that each leg is the bigger than min volume
        For i = 1 To NumLegs
            If m_oCurrentVeh.Components(legarray(i)).Volume < MinLegVolume Then
                retval = "Each leg must be greater than " & MinLegVolume & " cf. of volume.  Add more legs to balance the load or add empty space to the legs to increase their volume."
            End If
        Next
        ' Check to see that each leg is also the same volume
        For j = 2 To NumLegs
            If arrLegVolumes(1) <> m_oCurrentVeh.Components(legarray(j)).Volume Then
                retval2 = "Each leg must be the same volume as the other legs.  Check volumes of each leg and add empty space if necessary"
            End If
        Next
    End If
    
    If retval = "" Then
        retval = retval2
    Else
        retval = retval & " " & retval2
    End If
    LegCheck = retval
    Exit Function
    
errorhandler:
    
End Function

Function HardpointCapacityCheck() As String
    Dim retval As String
    Dim element As Object
    On Error GoTo errorhandler
    Dim sParentKey As String
    
    For Each element In m_oCurrentVeh.Components
        sParentKey = element.LogicalParent
        If sParentKey <> "" Then
            If (TypeOf m_oCurrentVeh.Components(sParentKey) Is clsHardPoint) Then
                If element.Weight > m_oCurrentVeh.Components(sParentKey).loadcapacity Then
                    retval = retval & element.Description & " weight exceeds the maximum capacity of the hardpoint or weaponbay its attached to. "
                End If
            End If
        End If
    Next
    HardpointCapacityCheck = retval
    Exit Function
    
errorhandler:
    
    
    
End Function


Function SolarCellArrayCheck() As String
    Dim retval As String
    Dim element As Object
    On Error GoTo errorhandler
    Dim sParentKey As String
    Dim bStealth As Boolean
    Dim bInfrared As Boolean
    Dim bChameleon As Boolean
    Dim bLiquidCrystal As Boolean
    Dim sTemp As String
    Dim TL As Long
    Dim bNotOnPanels As Boolean
    Dim bCells As Boolean
    
    ' check for solar cells on vehicles with stealth, infrared cloaking, chameleon or liquid crystal skins
    With m_oCurrentVeh.surface
        TL = m_oCurrentVeh.Components(BODY_KEY).TL
        If .infraredcloaking <> "none" Then bInfrared = True
        If .Stealth <> "none" Then bStealth = True
        If .Chameleon <> "none" Then bChameleon = True
        If Not .LiquidCrystal Then bLiquidCrystal = True
    End With
        
    ' check that only TL10+ can have solar cells on gasbags
    For Each element In m_oCurrentVeh.Components
        If TypeOf element Is clsSolarCellArray Then
            bCells = True
            sParentKey = element.LogicalParent
            If sParentKey <> "" Then
                If (TypeOf m_oCurrentVeh.Components(sParentKey) Is clsGasbag) Then
                    If TL < 10 Then
                        retval = retval & element.Description & " should not be placed on gasbags prior to TL 10. "
                        bNotOnPanels = True
                        Exit For
                    End If
                ElseIf Not TypeOf m_oCurrentVeh.Components(sParentKey) Is clsSolarPanel Then
                    bNotOnPanels = True
                End If
            End If
        End If
    Next
    
    If (bCells) And (TL < 11) And (bNotOnPanels) Then
        If (bInfrared) Or (bStealth) Or (bChameleon) Or (bLiquidCrystal) Then
            If retval <> "" Then
                retval = retval & vbNewLine
            End If
        
            retval = retval & "Vehicles with stealth, infrared cloaking, chameleon systems or liquid crystal skins should not have solar cells on them prior to TL11 unless they are on solar panels."

        End If
    End If
    
    SolarCellArrayCheck = retval
    Exit Function
    
errorhandler:
    
    
    
End Function

Function TopSpeedRestrictions() As String
    Dim pType As String
    Dim element As Object
    Dim sRet As String
    
    For Each element In m_oCurrentVeh.PerformanceProfiles
        
            Select Case element.Datatype
                Case PERFORMANCEHOVER, PERFORMANCEAIR, PERFORMANCEMAGLEV
                    If element.DesignCheckString <> "" Then
                    
                        sRet = sRet & "FYI: In your performance profile named '" & element.Key & "' " & element.DesignCheckString
                    End If
            End Select
            

    Next

    TopSpeedRestrictions = sRet

End Function
Function SoftwareComplexityCheck() As String
    '//check to make sure that software is not of greater complexity
    '//then the computer its added to
    Dim element As Object
    Dim retval As String
    Dim sParent As String
    
    On Error GoTo errorhandler
    
    For Each element In m_oCurrentVeh.Components
        sParent = element.LogicalParent
        If sParent <> "" Then
            If TypeOf m_oCurrentVeh.Components(sParent) Is clsComputer Then
                If element.complexity > m_oCurrentVeh.Components(sParent).complexity Then
                    retval = retval & element.Description & " complexity exceeds the complexity rating of the " & element.Description & " its attached to. "
                End If
            End If
        End If
    Next
    SoftwareComplexityCheck = retval
    Exit Function
    
errorhandler:
    
End Function

''///////////////////////////////////////////////////////////////////////////
'' TODO: These have been ripped out of old classes and pasted here so i dont forget that these need
''       to be added as design checks and not hard coded!!!
''Public Property Let Stealth(ByVal b As Byte)
''Dim element As Object
''If b = RADICAL Then
''    'todo: this would be design check! not here
''        'Radical stealth cannot be added to vehicles with OpenMounts
''        ' Triplane or Biplane wings, masts, or exposed or cycle seats
'''    For Each element In Veh.Components
'''
'''        If TypeOf element Is clsMast Then
'''            MsgBox "Radical Stealth cannot be applied to Vehicles with Masts"
'''            Exit Property
'''        'check for exposed Roomy, CrampedSeat or NormalSeat
'''        ElseIf TypeOf element Is clsAccommodation Then
'''            If (element.Datatype = RoomySeat) Or (element.Datatype = CrampedSeat) Or (element.Datatype = NormalSeat) And (element.Exposed) Then
'''                MsgBox "Radical Stealth cannot be applied to Vehicles with Exposed Seats"
'''                Exit Property
'''            ElseIf element.Datatype = CycleSeat Then
'''                MsgBox "Radical Stealth cannot be applied to Vehicles with Cycle Seats"
'''                Exit Property
'''            End If
'''        ElseIf TypeOf element Is clsWing Then
'''            If element.SubType = "biplane" Then
'''                MsgBox "Radical Stealth cannot be applied to Vehicles with Biplane Wings"
'''                Exit Property
'''            ElseIf element.SubType = "triplane" Then
'''                MsgBox "Radical Stealth cannot be applied to Vehicles with Triplane Wings"
'''                Exit Property
'''            End If
'''        End If
'''    Next
''End If
''    ' we havent exited the sub prematurely so everything checks out
''    m_byteStealth = b
''End Property
''
''Public Property Let EmissionCloaking(ByVal b As Byte)
''    If m_byteEmissionCloaking <> 0 Then
''        'if emission cloaking is anything except NONE, then infared cloaking has to be none
''        ' since it also covers infrared automatically
''        'todo: this should be a design check and not hardcoded here!
''        m_byteInfraredCloaking = 0
''    End If
''    m_byteEmissionCloaking = b
''End Property
''Public Property Let Sealed(ByVal bln As Boolean)
''    'todo: a design check! not a hard coded prevention
''    If mvarFloatationHull Then
''        If (bln = False) And (mvarWaterProof = False) Then
''            InfoPrint 1, "A vehicle with a floation hull must be sealed OR waterproofed. (Note: Sealed vehicles are waterproofed for free.)"
''        End If
''    End If
''    ' sealed also includes free waterproofing
''    If vdata = True Then mvarWaterProof = True
''    mvarSealed = bln
''End Property
''Public Property Let WaterProof(ByVal bln As Boolean)
''    'todo: a design check! not a hard coded prevention
''    If mvarFloatationHull Then
''        If (bln = False) And (mvarSealed = False) Then
''            InfoPrint 1, "A vehicle with a floation hull must be sealed OR waterproofed. (Note: Sealed vehicles are waterproofed for free.)"
''        End If
''    End If
''    mvarWaterProof = bln
''End Property
''Public Property Let HydrodynamicLines(ByVal s As String)
''    If s = "submarine" Then
''        mvarSubmersible = True
''        ' todo: this is a design check for the hydroynamiclines "submersible" option
''        InfoPrint 1, "A vehicle with submarine lines must have Submersible Option enabled. This has been done for you."
''    End If
''    mvarHydrodynamicLines = s
''End Property
''Public Property Let StreamLining(ByVal s As String)
''    If veh.Components(BODY_KEY).LiftingBody Then
''        Select Case s
''            'todo: design check
''            Case "fair", "good", "none"  'fixed. Fair and Good were capitalized and they shouldn't be.  This is why we need to switch to constants! MPJ 01/27/2004
''                InfoPrint 1, "Your Vehicle has the 'lifting body' option enabled.  This requires 'very good' or better streamlining."
''                s = "very good"
''                Exit Property
''        End Select
''    End If
''    mvarStreamLining = s
''End Property
''Public Property Let Submersible(ByVal bln As Boolean)
''    ' if submersible, the vehicle must be sealed and waterproofed.  Its free
''    ' however
''    If bln = True Then
''        mvarWaterProof = True
''        mvarSealed = True
''    End If
''    mvarSubmersible = bln
''End Property
