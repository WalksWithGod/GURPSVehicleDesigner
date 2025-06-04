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
vbwProfiler.vbwProcIn 107
vbwProfiler.vbwExecuteLine 2761
    Select Case KeyCode
'vbwLine 2762:        Case vbKeyEscape
        Case IIf(vbwProfiler.vbwExecuteLine(2762), VBWPROFILER_EMPTY, _
        vbKeyEscape)
vbwProfiler.vbwExecuteLine 2763
            Unload Me
    End Select
vbwProfiler.vbwExecuteLine 2764 'B
vbwProfiler.vbwProcOut 107
vbwProfiler.vbwExecuteLine 2765
End Sub

Private Sub Form_Load()
vbwProfiler.vbwProcIn 108
    Dim sRet As String
    Dim sTotal As String

    'TODO: this was removed from NewProfile and needs to be added somewhere
    'if user chose to make Submerged, check to make sure Vehicle Options Submerged is true
    'If cboPerformanceType.Text = "Submerged" And m_oCurrentVeh.Components(BODY_KEY).Submersible <> True Then
    '    MsgBox "You have not yet enabled the Submersible property in the Options dialog.  This will be done for you."
    '    m_oCurrentVeh.Components(BODY_KEY).Submersible = True
    'End If


vbwProfiler.vbwExecuteLine 2766
    sTotal = ""
vbwProfiler.vbwExecuteLine 2767
    sTotal = MinBodyVolumeCheck
vbwProfiler.vbwExecuteLine 2768
    MsgBox "Design Check not yet implemented..."
vbwProfiler.vbwProcOut 108
vbwProfiler.vbwExecuteLine 2769
    Exit Sub
    ' check that weights loaded onto hardpoitns dont exceed the hardpoints' capacity
vbwProfiler.vbwExecuteLine 2770
    sRet = HardpointCapacityCheck
vbwProfiler.vbwExecuteLine 2771
    Call AppendCheckText(sTotal, sRet)

    ' check leg volumes match and that they meet minimum volume reqts
vbwProfiler.vbwExecuteLine 2772
    sRet = LegCheck
vbwProfiler.vbwExecuteLine 2773
    Call AppendCheckText(sTotal, sRet)


vbwProfiler.vbwExecuteLine 2774
    sRet = SoftwareComplexityCheck
vbwProfiler.vbwExecuteLine 2775
    Call AppendCheckText(sTotal, sRet)

    ' check that if a duplicate maneuver control is found, that an original (non duplicate) one exists as well
vbwProfiler.vbwExecuteLine 2776
    sRet = DuplicateControlsCheck
vbwProfiler.vbwExecuteLine 2777
    Call AppendCheckText(sTotal, sRet)

    ' check that all subs with have armor if streamlining is used
vbwProfiler.vbwExecuteLine 2778
    sRet = ArmorCheck
vbwProfiler.vbwExecuteLine 2779
    Call AppendCheckText(sTotal, sRet)

    ' check max non rigid armor DR's are not exceeded
vbwProfiler.vbwExecuteLine 2780
    sRet = MaxNonRigidArmorDRCheck
vbwProfiler.vbwExecuteLine 2781
    Call AppendCheckText(sTotal, sRet)

    ' check that engines are connected to fuel tanks which use the same type of fuel
vbwProfiler.vbwExecuteLine 2782
    sRet = EngineTankFuelTypeCheck
vbwProfiler.vbwExecuteLine 2783
    Call AppendCheckText(sTotal, sRet)

    ' check for hardpoints on body exception rules
vbwProfiler.vbwExecuteLine 2784
    sRet = HardpointsOnBodyCheck
vbwProfiler.vbwExecuteLine 2785
    Call AppendCheckText(sTotal, sRet)

    ' check for streamlining rules violations
vbwProfiler.vbwExecuteLine 2786
    sRet = StreamliningViolationsCheck
vbwProfiler.vbwExecuteLine 2787
    Call AppendCheckText(sTotal, sRet)

    ' check for recommended minimum wing volumes
vbwProfiler.vbwExecuteLine 2788
    sRet = WingVolumesCheck
vbwProfiler.vbwExecuteLine 2789
    Call AppendCheckText(sTotal, sRet)

    ' check for solar cell placement violations
vbwProfiler.vbwExecuteLine 2790
    sRet = SolarCellArrayCheck
vbwProfiler.vbwExecuteLine 2791
    Call AppendCheckText(sTotal, sRet)

    ' inform user for reasons of any speed limits they may have hit
vbwProfiler.vbwExecuteLine 2792
    sRet = TopSpeedRestrictions
vbwProfiler.vbwExecuteLine 2793
    Call AppendCheckText(sTotal, sRet)

vbwProfiler.vbwExecuteLine 2794
    sRet = StackedTurretsCheck
vbwProfiler.vbwExecuteLine 2795
    Call AppendCheckText(sTotal, sRet)

    ' Design Check:  Make sure vehicles Max Lift or Floatation rating has not been exceeded
' TODO If it does tell user to convert cargo space to empty space
' to reduce its weight or eliminate some armor.
vbwProfiler.vbwExecuteLine 2796
    txtDesignCheck = sTotal

vbwProfiler.vbwExecuteLine 2797
    If txtDesignCheck = "" Then
vbwProfiler.vbwExecuteLine 2798
         txtDesignCheck = "No design flaws detected in current vehicle."
    End If
vbwProfiler.vbwExecuteLine 2799 'B

vbwProfiler.vbwProcOut 108
vbwProfiler.vbwExecuteLine 2800
End Sub

Sub AppendCheckText(ByRef sTarget As String, ByRef sChunk As String)
vbwProfiler.vbwProcIn 109

vbwProfiler.vbwExecuteLine 2801
    If sChunk <> "" Then

vbwProfiler.vbwExecuteLine 2802
        If sTarget = "" Then
vbwProfiler.vbwExecuteLine 2803
            sTarget = sChunk
        Else
vbwProfiler.vbwExecuteLine 2804 'B
vbwProfiler.vbwExecuteLine 2805
            sTarget = sTarget & vbNewLine & sChunk
        End If
vbwProfiler.vbwExecuteLine 2806 'B
    End If
vbwProfiler.vbwExecuteLine 2807 'B


vbwProfiler.vbwProcOut 109
vbwProfiler.vbwExecuteLine 2808
End Sub

Function StackedTurretsCheck() As String
vbwProfiler.vbwProcIn 110
    Dim sParent As String
    Dim Temp As String
    Dim subsarray() As String
    Dim i, num As Long
    Dim sRet As String

vbwProfiler.vbwExecuteLine 2809
    On Error Resume Next

    ' get all our subassemblies.  If we have none except the body, exit function
vbwProfiler.vbwExecuteLine 2810
    subsarray = m_oCurrentVeh.Components(BODY_KEY).GetCurrentSubAssembliesKeys
vbwProfiler.vbwExecuteLine 2811
    num = UBound(subsarray)
vbwProfiler.vbwExecuteLine 2812
    If subsarray(1) = "" Then
vbwProfiler.vbwProcOut 110
vbwProfiler.vbwExecuteLine 2813
         Exit Function
    End If
vbwProfiler.vbwExecuteLine 2814 'B

    ' MPJ 11/7/2000
    ' this has been removed as a hard restriction in the
    ' turrets Let Orientation() property and is
    ' place here in the design check instead
vbwProfiler.vbwExecuteLine 2815
    For i = 1 To num
vbwProfiler.vbwExecuteLine 2816
        If TypeOf m_oCurrentVeh.Components(subsarray(i)) Is clsTurret Then
vbwProfiler.vbwExecuteLine 2817
            sParent = m_oCurrentVeh.Components(subsarray(i)).Parent  'note: NOT logical parent since subassemblies cant be attached to Groups
vbwProfiler.vbwExecuteLine 2818
            If TypeOf m_oCurrentVeh.Components(sParent) Is clsTurret Then

vbwProfiler.vbwExecuteLine 2819
                Temp = m_oCurrentVeh.Components(sParent).Orientation
vbwProfiler.vbwExecuteLine 2820
                If Temp <> m_oCurrentVeh.Components(subsarray(i)).Orientation Then
vbwProfiler.vbwExecuteLine 2821
                    sRet = "FYI: Stacked Turrets should usually share the same orientation."
                End If
vbwProfiler.vbwExecuteLine 2822 'B
            End If
vbwProfiler.vbwExecuteLine 2823 'B
        End If
vbwProfiler.vbwExecuteLine 2824 'B
vbwProfiler.vbwExecuteLine 2825
    Next

vbwProfiler.vbwExecuteLine 2826
    StackedTurretsCheck = sRet
vbwProfiler.vbwProcOut 110
vbwProfiler.vbwExecuteLine 2827
End Function

Function WingVolumesCheck() As String
    'checks that wings are at the recommended 0.1 x body volume and for stub wings at least 0.02 x body volume
vbwProfiler.vbwProcIn 111

    Dim element As Object
    Dim dblBodyVolume As Double
    Dim dblWingVolume As Double
    Dim sRet As String
    Dim dblMinStub As Double
    Dim dblMinStandard As Double

vbwProfiler.vbwExecuteLine 2828
    dblBodyVolume = m_oCurrentVeh.Components(BODY_KEY).Volume

vbwProfiler.vbwExecuteLine 2829
    dblMinStandard = Round(0.1 * dblBodyVolume, 2)
vbwProfiler.vbwExecuteLine 2830
    dblMinStub = Round(0.02 * dblBodyVolume, 2)

vbwProfiler.vbwExecuteLine 2831
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 2832
        If TypeOf element Is clsWing Then

vbwProfiler.vbwExecuteLine 2833
            dblWingVolume = element.Volume

vbwProfiler.vbwExecuteLine 2834
            If element.subtype = "stub" Then
vbwProfiler.vbwExecuteLine 2835
                If dblWingVolume < dblMinStub Then
vbwProfiler.vbwExecuteLine 2836
                    sRet = "WARNING: Detected a wing volume of " & dblWingVolume & ".  Minimum recommended volume for 'stub wings' is 0.02 x Body Volume. In your case, each stub wing should be at least " & dblMinStub & " cf."
                End If
vbwProfiler.vbwExecuteLine 2837 'B
            Else
vbwProfiler.vbwExecuteLine 2838 'B
vbwProfiler.vbwExecuteLine 2839
                If dblWingVolume < dblMinStandard Then
vbwProfiler.vbwExecuteLine 2840
                    sRet = "WARNING: Detected a wing volume of " & dblWingVolume & ".  Minimum recommended volume for wings is 0.1 x Body Volume. In your case, each wing should be at least " & dblMinStandard & " cf."
                End If
vbwProfiler.vbwExecuteLine 2841 'B
            End If
vbwProfiler.vbwExecuteLine 2842 'B
        End If
vbwProfiler.vbwExecuteLine 2843 'B
vbwProfiler.vbwExecuteLine 2844
    Next

vbwProfiler.vbwExecuteLine 2845
    WingVolumesCheck = sRet
vbwProfiler.vbwProcOut 111
vbwProfiler.vbwExecuteLine 2846
End Function


Function StreamliningViolationsCheck() As String
    'VE page 11, "a vehicle with masts cannot have better than Fair streamlining.
    ' A vehicle with skids or wheels (unless retractable), tracks, halftracks, skitraacks
    ' biplane or triaplane wings, GEV skirs, SEV sidewalls, open mounts, gasbags, rotors,
    ' arms or legs cannot have better than Good streamling.  A vehicle with superstructures
    ' or turrets (except pop turrets) cannot have better than Very Good streamling. Vehicles
    ' with wings cannot have Superior, Excellent or Radical streamlining before TL7.
vbwProfiler.vbwProcIn 112


    Dim sType As String
    Dim lngDType As Long
    Dim subsarray() As String
    Dim i, num As Long
    Dim sRet As String

vbwProfiler.vbwExecuteLine 2847
    On Error Resume Next

    ' get all our subassemblies.  If we have none except the body, exit function
vbwProfiler.vbwExecuteLine 2848
    subsarray = m_oCurrentVeh.Components(BODY_KEY).GetCurrentSubAssembliesKeys
vbwProfiler.vbwExecuteLine 2849
    num = UBound(subsarray)
vbwProfiler.vbwExecuteLine 2850
    If subsarray(1) = "" Then
vbwProfiler.vbwProcOut 112
vbwProfiler.vbwExecuteLine 2851
         Exit Function
    End If
vbwProfiler.vbwExecuteLine 2852 'B

    ' determine which level of streamlining the user has
vbwProfiler.vbwExecuteLine 2853
    sType = m_oCurrentVeh.surface.StreamLining

vbwProfiler.vbwExecuteLine 2854
    Select Case sType
'vbwLine 2855:        Case "none", "fair"
        Case IIf(vbwProfiler.vbwExecuteLine(2855), VBWPROFILER_EMPTY, _
        "none"), "fair"
vbwProfiler.vbwProcOut 112
vbwProfiler.vbwExecuteLine 2856
            Exit Function

'vbwLine 2857:        Case "good", "very good", "superior", "excellent", "radical"
        Case IIf(vbwProfiler.vbwExecuteLine(2857), VBWPROFILER_EMPTY, _
        "good"), "very good", "superior", "excellent", "radical"
            ' check for masts limit of "fair" streamlining
vbwProfiler.vbwExecuteLine 2858
            For i = 1 To num
vbwProfiler.vbwExecuteLine 2859
                lngDType = m_oCurrentVeh.Components(subsarray(i)).Datatype

vbwProfiler.vbwExecuteLine 2860
                Select Case lngDType
'vbwLine 2861:                    Case Mast
                    Case IIf(vbwProfiler.vbwExecuteLine(2861), VBWPROFILER_EMPTY, _
        Mast)
vbwProfiler.vbwExecuteLine 2862
                        sRet = "WARNING: You have Masts installed and " & sType & " streamlining. Vehicles with Masts cant have better than 'fair' streamlining."

                End Select
vbwProfiler.vbwExecuteLine 2863 'B
vbwProfiler.vbwExecuteLine 2864
            Next

            ' check for rules which limit streamlining to "good" or lower
vbwProfiler.vbwExecuteLine 2865
            If sType <> "good" Then
vbwProfiler.vbwExecuteLine 2866
                For i = 1 To num
vbwProfiler.vbwExecuteLine 2867
                    lngDType = m_oCurrentVeh.Components(subsarray(i)).Datatype

vbwProfiler.vbwExecuteLine 2868
                    Select Case lngDType

'vbwLine 2869:                        Case Skid ' unless retractable
                        Case IIf(vbwProfiler.vbwExecuteLine(2869), VBWPROFILER_EMPTY, _
        Skid )' unless retractable
vbwProfiler.vbwExecuteLine 2870
                            If (m_oCurrentVeh.Components(subsarray(i)).RetractLocation = "none") Then
vbwProfiler.vbwExecuteLine 2871
                                If sRet <> "" Then
vbwProfiler.vbwExecuteLine 2872
                                     sRet = sRet & vbNewLine
                                End If
vbwProfiler.vbwExecuteLine 2873 'B
vbwProfiler.vbwExecuteLine 2874
                                sRet = sRet & "WARNING: Your vehicle has non-retractable skids installed and " & sType & " streamlining.  Vehicles with with non-retractable skids cannot have better than 'good' streamlining."
                            End If
vbwProfiler.vbwExecuteLine 2875 'B

'vbwLine 2876:                        Case Wheel 'unless retractable
                        Case IIf(vbwProfiler.vbwExecuteLine(2876), VBWPROFILER_EMPTY, _
        Wheel )'unless retractable
vbwProfiler.vbwExecuteLine 2877
                            If (m_oCurrentVeh.Components(subsarray(i)).subtype <> "retractable") Then
vbwProfiler.vbwExecuteLine 2878
                                If sRet <> "" Then
vbwProfiler.vbwExecuteLine 2879
                                     sRet = sRet & vbNewLine
                                End If
vbwProfiler.vbwExecuteLine 2880 'B
vbwProfiler.vbwExecuteLine 2881
                                sRet = sRet & "WARNING: Your vehicle has non-retractable wheels installed and " & sType & " streamlining.  Vehicles with with non-retractable wheels cannot have better than 'good' streamlining."
                            End If
vbwProfiler.vbwExecuteLine 2882 'B

'vbwLine 2883:                        Case Track
                        Case IIf(vbwProfiler.vbwExecuteLine(2883), VBWPROFILER_EMPTY, _
        Track)
vbwProfiler.vbwExecuteLine 2884
                            If sRet <> "" Then
vbwProfiler.vbwExecuteLine 2885
                                 sRet = sRet & vbNewLine
                            End If
vbwProfiler.vbwExecuteLine 2886 'B
vbwProfiler.vbwExecuteLine 2887
                            sRet = sRet & "WARNING: Your vehicle has a Track subassembly installed and " & sType & " streamlining.  Vehicles with with Tracks cannot have better than 'good' streamlining."


'vbwLine 2888:                        Case Wing ' biplane or triaplane wings only
                        Case IIf(vbwProfiler.vbwExecuteLine(2888), VBWPROFILER_EMPTY, _
        Wing )' biplane or triaplane wings only
vbwProfiler.vbwExecuteLine 2889
                            If (m_oCurrentVeh.Components(subsarray(i)).subtype = "biplane") Or (m_oCurrentVeh.Components(subsarray(i)).subtype = "triplane") Then
vbwProfiler.vbwExecuteLine 2890
                                If sRet <> "" Then
vbwProfiler.vbwExecuteLine 2891
                                     sRet = sRet & vbNewLine
                                End If
vbwProfiler.vbwExecuteLine 2892 'B
vbwProfiler.vbwExecuteLine 2893
                                sRet = sRet & "WARNING: Your vehicle has biplane or triplane wings installed and " & sType & " streamlining.  Vehicles with bi/triplane wings cannot have better than 'good' streamlining."
                            End If
vbwProfiler.vbwExecuteLine 2894 'B
'vbwLine 2895:                        Case Hovercraft
                        Case IIf(vbwProfiler.vbwExecuteLine(2895), VBWPROFILER_EMPTY, _
        Hovercraft)
vbwProfiler.vbwExecuteLine 2896
                            If sRet <> "" Then
vbwProfiler.vbwExecuteLine 2897
                                 sRet = sRet & vbNewLine
                            End If
vbwProfiler.vbwExecuteLine 2898 'B
vbwProfiler.vbwExecuteLine 2899
                            sRet = sRet & "WARNING: Your vehicle has a GEV or SEV hovercraft subassembly installed and " & sType & " streamlining.  Vehicles with with GEV or SEV cannot have better than 'good' streamlining."

'vbwLine 2900:                        Case OpenMount
                        Case IIf(vbwProfiler.vbwExecuteLine(2900), VBWPROFILER_EMPTY, _
        OpenMount)
vbwProfiler.vbwExecuteLine 2901
                            If sRet <> "" Then
vbwProfiler.vbwExecuteLine 2902
                                 sRet = sRet & vbNewLine
                            End If
vbwProfiler.vbwExecuteLine 2903 'B
vbwProfiler.vbwExecuteLine 2904
                            sRet = sRet & "WARNING: Your vehicle has Open Mount subassemblies installed and " & sType & " streamlining.  Vehicles with with Open Mounts cannot have better than 'good' streamlining."

'vbwLine 2905:                        Case Gasbag
                        Case IIf(vbwProfiler.vbwExecuteLine(2905), VBWPROFILER_EMPTY, _
        Gasbag)
vbwProfiler.vbwExecuteLine 2906
                            If sRet <> "" Then
vbwProfiler.vbwExecuteLine 2907
                                 sRet = sRet & vbNewLine
                            End If
vbwProfiler.vbwExecuteLine 2908 'B
vbwProfiler.vbwExecuteLine 2909
                            sRet = sRet & "WARNING: Your vehicle has a Gasbag subassembly installed and " & sType & " streamlining.  Vehicles with with Gasbags cannot have better than 'good' streamlining."

'vbwLine 2910:                        Case AutogyroRotor, TTRotor, CARotor, MMRotor
                        Case IIf(vbwProfiler.vbwExecuteLine(2910), VBWPROFILER_EMPTY, _
        AutogyroRotor), TTRotor, CARotor, MMRotor
vbwProfiler.vbwExecuteLine 2911
                            If sRet <> "" Then
vbwProfiler.vbwExecuteLine 2912
                                 sRet = sRet & vbNewLine
                            End If
vbwProfiler.vbwExecuteLine 2913 'B
vbwProfiler.vbwExecuteLine 2914
                            sRet = sRet & "WARNING: Your vehicle has a Rotor subassembly installed and " & sType & " streamlining.  Vehicles with with Rotors cannot have better than 'good' streamlining."

'vbwLine 2915:                        Case Arm
                        Case IIf(vbwProfiler.vbwExecuteLine(2915), VBWPROFILER_EMPTY, _
        Arm)
vbwProfiler.vbwExecuteLine 2916
                             If sRet <> "" Then
vbwProfiler.vbwExecuteLine 2917
                                  sRet = sRet & vbNewLine
                             End If
vbwProfiler.vbwExecuteLine 2918 'B
vbwProfiler.vbwExecuteLine 2919
                            sRet = sRet & "WARNING: Your vehicle has an Arm subassembly installed and " & sType & " streamlining.  Vehicles with with Arms cannot have better than 'good' streamlining."

'vbwLine 2920:                        Case Leg
                        Case IIf(vbwProfiler.vbwExecuteLine(2920), VBWPROFILER_EMPTY, _
        Leg)
vbwProfiler.vbwExecuteLine 2921
                             If sRet <> "" Then
vbwProfiler.vbwExecuteLine 2922
                                  sRet = sRet & vbNewLine
                             End If
vbwProfiler.vbwExecuteLine 2923 'B
vbwProfiler.vbwExecuteLine 2924
                            sRet = sRet & "WARNING: Your vehicle has a Leg subassembly installed and " & sType & " streamlining.  Vehicles with with Legs cannot have better than 'good' streamlining."


                    End Select
vbwProfiler.vbwExecuteLine 2925 'B
vbwProfiler.vbwExecuteLine 2926
                Next

            ' check for rules which limit streamlining to "very good" or lower
'vbwLine 2927:            ElseIf sType <> "very good" Then
            ElseIf vbwProfiler.vbwExecuteLine(2927) Or sType <> "very good" Then
vbwProfiler.vbwExecuteLine 2928
                For i = 1 To num
vbwProfiler.vbwExecuteLine 2929
                    lngDType = m_oCurrentVeh.Components(subsarray(i)).Datatype

vbwProfiler.vbwExecuteLine 2930
                    Select Case lngDType
'vbwLine 2931:                        Case Turret
                        Case IIf(vbwProfiler.vbwExecuteLine(2931), VBWPROFILER_EMPTY, _
        Turret)
vbwProfiler.vbwExecuteLine 2932
                            If sRet <> "" Then
vbwProfiler.vbwExecuteLine 2933
                                 sRet = sRet & vbNewLine
                            End If
vbwProfiler.vbwExecuteLine 2934 'B
vbwProfiler.vbwExecuteLine 2935
                            sRet = sRet & "WARNING: You have turrets installed and " & sType & " streamlining.  Vehicles with turrets cannot have better than "

'vbwLine 2936:                        Case Superstructure
                        Case IIf(vbwProfiler.vbwExecuteLine(2936), VBWPROFILER_EMPTY, _
        Superstructure)
vbwProfiler.vbwExecuteLine 2937
                            If sRet <> "" Then
vbwProfiler.vbwExecuteLine 2938
                                 sRet = sRet & vbNewLine
                            End If
vbwProfiler.vbwExecuteLine 2939 'B
vbwProfiler.vbwExecuteLine 2940
                            sRet = sRet & "WARNING: You have superstructures installed and " & sType & " streamlining.  Vehicle with superstructures cannot have better than 'very good' streamlining."

'vbwLine 2941:                        Case Wing
                        Case IIf(vbwProfiler.vbwExecuteLine(2941), VBWPROFILER_EMPTY, _
        Wing)
vbwProfiler.vbwExecuteLine 2942
                            If m_oCurrentVeh.Components(BODY_KEY).TL < 7 Then
                            ' if it has wings and before TL 7, then it cannot have superior or better streamlining

vbwProfiler.vbwExecuteLine 2943
                                If sRet <> "" Then
vbwProfiler.vbwExecuteLine 2944
                                     sRet = sRet & vbNewLine
                                End If
vbwProfiler.vbwExecuteLine 2945 'B
vbwProfiler.vbwExecuteLine 2946
                                sRet = sRet & "WARNING: You have wings installed and " & sType & " streamlining.  Vehicles prior to TL7 cannot have better than 'very good' streamlining."
                            End If
vbwProfiler.vbwExecuteLine 2947 'B
                    End Select
vbwProfiler.vbwExecuteLine 2948 'B
vbwProfiler.vbwExecuteLine 2949
                Next
            End If
vbwProfiler.vbwExecuteLine 2950 'B

    End Select
vbwProfiler.vbwExecuteLine 2951 'B

vbwProfiler.vbwExecuteLine 2952
    StreamliningViolationsCheck = sRet

vbwProfiler.vbwProcOut 112
vbwProfiler.vbwExecuteLine 2953
End Function

Function HardpointsOnBodyCheck() As String
    'page 94 - hardpoints may not be added to a vehicle's body if it has a hydrodynamic hull,
    'GEV or SEV subassemblies, tracks, halftracks or skitracks, railway wheels or a
    'flexibody drivetrain.
vbwProfiler.vbwProcIn 113

    Dim element As Object
    Dim bHydro As Boolean
    Dim bGEVSEV As Boolean
    Dim bTracks As Boolean
    Dim bRailwayWheels As Boolean
    Dim bFlexibody As Boolean
    Dim bHardpointsOnBody As Boolean
    Dim sRet As String

vbwProfiler.vbwExecuteLine 2954
    On Error Resume Next

    ' determine if we have hardpoints on the body
vbwProfiler.vbwExecuteLine 2955
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 2956
        If TypeOf element Is clsHardPoint Then
vbwProfiler.vbwExecuteLine 2957
            If element.Datatype = HardPoint Then
vbwProfiler.vbwExecuteLine 2958
                bHardpointsOnBody = True
vbwProfiler.vbwExecuteLine 2959
                Exit For
            End If
vbwProfiler.vbwExecuteLine 2960 'B
        End If
vbwProfiler.vbwExecuteLine 2961 'B
vbwProfiler.vbwExecuteLine 2962
    Next

    ' if we do have hardpoints on body, check to see if exceptions are found which
    ' recommend hardpoints not be added to body
vbwProfiler.vbwExecuteLine 2963
    If bHardpointsOnBody Then
vbwProfiler.vbwExecuteLine 2964
        For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 2965
            If TypeOf element Is clsBody Then
vbwProfiler.vbwExecuteLine 2966
                If element.HydrodynamicLines <> "none" Then
vbwProfiler.vbwExecuteLine 2967
                    bHydro = True
                End If
vbwProfiler.vbwExecuteLine 2968 'B
'vbwLine 2969:            ElseIf TypeOf element Is clsHovercraft Then
            ElseIf vbwProfiler.vbwExecuteLine(2969) Or TypeOf element Is clsHovercraft Then
vbwProfiler.vbwExecuteLine 2970
                bGEVSEV = True
'vbwLine 2971:            ElseIf TypeOf element Is clsTrack Then
            ElseIf vbwProfiler.vbwExecuteLine(2971) Or TypeOf element Is clsTrack Then
vbwProfiler.vbwExecuteLine 2972
                bTracks = True
'vbwLine 2973:            ElseIf TypeOf element Is clsWheel Then
            ElseIf vbwProfiler.vbwExecuteLine(2973) Or TypeOf element Is clsWheel Then
vbwProfiler.vbwExecuteLine 2974
                If element.subtype = "railway" Then
vbwProfiler.vbwExecuteLine 2975
                    bRailwayWheels = True
                End If
vbwProfiler.vbwExecuteLine 2976 'B
'vbwLine 2977:            ElseIf TypeOf element Is clsGroundDrivetrain Then
            ElseIf vbwProfiler.vbwExecuteLine(2977) Or TypeOf element Is clsGroundDrivetrain Then
vbwProfiler.vbwExecuteLine 2978
                If element.Datatype = FlexibodyDrivetrain Then
vbwProfiler.vbwExecuteLine 2979
                    bFlexibody = True
                End If
vbwProfiler.vbwExecuteLine 2980 'B
            End If
vbwProfiler.vbwExecuteLine 2981 'B
vbwProfiler.vbwExecuteLine 2982
        Next
    Else
vbwProfiler.vbwExecuteLine 2983 'B
vbwProfiler.vbwProcOut 113
vbwProfiler.vbwExecuteLine 2984
        Exit Function
    End If
vbwProfiler.vbwExecuteLine 2985 'B

    ' create print output
vbwProfiler.vbwExecuteLine 2986
    If bHydro Then
vbwProfiler.vbwExecuteLine 2987
         sRet = sRet & "hydrodynamic lines, "
    End If
vbwProfiler.vbwExecuteLine 2988 'B
vbwProfiler.vbwExecuteLine 2989
    If bGEVSEV Then
vbwProfiler.vbwExecuteLine 2990
         sRet = sRet & "a GEV or SEV hovercraft subassembly, "
    End If
vbwProfiler.vbwExecuteLine 2991 'B
vbwProfiler.vbwExecuteLine 2992
    If bTracks Then
vbwProfiler.vbwExecuteLine 2993
         sRet = sRet & "tracks, "
    End If
vbwProfiler.vbwExecuteLine 2994 'B
vbwProfiler.vbwExecuteLine 2995
    If bRailwayWheels Then
vbwProfiler.vbwExecuteLine 2996
         sRet = sRet & "railway wheels, "
    End If
vbwProfiler.vbwExecuteLine 2997 'B
vbwProfiler.vbwExecuteLine 2998
    If bFlexibody Then
vbwProfiler.vbwExecuteLine 2999
         sRet = sRet & "a flexibody drivetrain "
    End If
vbwProfiler.vbwExecuteLine 3000 'B

vbwProfiler.vbwExecuteLine 3001
    If sRet <> "" Then
vbwProfiler.vbwExecuteLine 3002
        sRet = "WARNING: You have hardpoints attached to the body. You also have " & sRet
vbwProfiler.vbwExecuteLine 3003
        sRet = sRet & ". VE 2nd edition suggests that hardpoints not be added to the body IF it has hydrodynamic lines " _
            & "GEV or SEV subassemblies, tracks, halftracks, or skitracks, railways wheels or a flexibody drivetrain."
    End If
vbwProfiler.vbwExecuteLine 3004 'B

vbwProfiler.vbwExecuteLine 3005
    HardpointsOnBodyCheck = sRet
vbwProfiler.vbwProcOut 113
vbwProfiler.vbwExecuteLine 3006
End Function

Function EngineTankFuelTypeCheck() As String
vbwProfiler.vbwProcIn 114

    Dim arrEngines() As String
    Dim arrTanks() As String
    Dim i As Long
    Dim j As Long
    Dim sText As String
    Dim sFuelType As String
    Dim sTankDescription As String
    Dim sEngineDescription As String

vbwProfiler.vbwExecuteLine 3007
    On Error Resume Next

vbwProfiler.vbwExecuteLine 3008
    arrTanks = m_oCurrentVeh.Components(BODY_KEY).GetCurrentFuelStorageKeys
vbwProfiler.vbwExecuteLine 3009
    If arrTanks(1) = "" Then
vbwProfiler.vbwProcOut 114
vbwProfiler.vbwExecuteLine 3010
         Exit Function
    End If
vbwProfiler.vbwExecuteLine 3011 'B

vbwProfiler.vbwExecuteLine 3012
    For i = 1 To UBound(arrTanks)
vbwProfiler.vbwExecuteLine 3013
        sFuelType = m_oCurrentVeh.Components(arrTanks(i)).Fuel
vbwProfiler.vbwExecuteLine 3014
        sTankDescription = m_oCurrentVeh.Components(arrTanks(i)).CustomDescription

vbwProfiler.vbwExecuteLine 3015
        arrEngines = m_oCurrentVeh.Components(BODY_KEY).GetCurrentFuelUsingSystemKeys
vbwProfiler.vbwExecuteLine 3016
        If arrEngines(1) <> "" Then
vbwProfiler.vbwExecuteLine 3017
            For j = 1 To UBound(arrEngines)
vbwProfiler.vbwExecuteLine 3018
                sEngineDescription = m_oCurrentVeh.Components(arrEngines(j)).CustomDescription
vbwProfiler.vbwExecuteLine 3019
                If m_oCurrentVeh.Components(arrEngines(j)).Fueltype <> sFuelType Then
vbwProfiler.vbwExecuteLine 3020
                    sText = sText & "Fuel type for " & sEngineDescription & " does not match fuel stored in " & sTankDescription & ".  "
                End If
vbwProfiler.vbwExecuteLine 3021 'B
vbwProfiler.vbwExecuteLine 3022
            Next
        End If
vbwProfiler.vbwExecuteLine 3023 'B
vbwProfiler.vbwExecuteLine 3024
    Next


vbwProfiler.vbwExecuteLine 3025
    EngineTankFuelTypeCheck = sText

vbwProfiler.vbwProcOut 114
vbwProfiler.vbwExecuteLine 3026
End Function

Function MaxNonRigidArmorDRCheck() As String
vbwProfiler.vbwProcIn 115

    Dim element As Object
    Dim dType As Long
    Dim sText As String
    Dim bReflex As Boolean
    Dim bFlag As Boolean

vbwProfiler.vbwExecuteLine 3027
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 3028
        If TypeOf element Is clsArmor Then
vbwProfiler.vbwExecuteLine 3029
            dType = element.Datatype

vbwProfiler.vbwExecuteLine 3030
            Select Case dType

'vbwLine 3031:                Case ArmorBasicFacing, ArmorOverall, ArmorWheelGuard, ArmorGunShield, ArmorOpenFrame, ArmorLocation, ArmorComponent
                Case IIf(vbwProfiler.vbwExecuteLine(3031), VBWPROFILER_EMPTY, _
        ArmorBasicFacing), ArmorOverall, ArmorWheelGuard, ArmorGunShield, ArmorOpenFrame, ArmorLocation, ArmorComponent
vbwProfiler.vbwExecuteLine 3032
                    If element.material = "nonrigid" Then
vbwProfiler.vbwExecuteLine 3033
                        bReflex = True
vbwProfiler.vbwExecuteLine 3034
                        If element.dr > 100 Then
vbwProfiler.vbwExecuteLine 3035
                            bFlag = True ' bail out as soon as we violate the rule
vbwProfiler.vbwExecuteLine 3036
                            Exit For
                        End If
vbwProfiler.vbwExecuteLine 3037 'B
                    End If
vbwProfiler.vbwExecuteLine 3038 'B

'vbwLine 3039:                Case ArmorComplexFacing
                Case IIf(vbwProfiler.vbwExecuteLine(3039), VBWPROFILER_EMPTY, _
        ArmorComplexFacing)
vbwProfiler.vbwExecuteLine 3040
                    If element.material1 = "nonrigid" Then
vbwProfiler.vbwExecuteLine 3041
                         bReflex = True
vbwProfiler.vbwExecuteLine 3042
                        If element.dr1 > 100 Then
vbwProfiler.vbwExecuteLine 3043
                            bFlag = True
vbwProfiler.vbwExecuteLine 3044
                            Exit For
                        End If
vbwProfiler.vbwExecuteLine 3045 'B
'vbwLine 3046:                    ElseIf element.material2 = "nonrigid" Then
                    ElseIf vbwProfiler.vbwExecuteLine(3046) Or element.material2 = "nonrigid" Then
vbwProfiler.vbwExecuteLine 3047
                         bReflex = True
vbwProfiler.vbwExecuteLine 3048
                        If element.dr2 > 100 Then
vbwProfiler.vbwExecuteLine 3049
                            bFlag = True
vbwProfiler.vbwExecuteLine 3050
                            Exit For
                        End If
vbwProfiler.vbwExecuteLine 3051 'B
'vbwLine 3052:                    ElseIf element.material3 = "nonrigid" Then
                    ElseIf vbwProfiler.vbwExecuteLine(3052) Or element.material3 = "nonrigid" Then
vbwProfiler.vbwExecuteLine 3053
                         bReflex = True
vbwProfiler.vbwExecuteLine 3054
                        If element.dr3 > 100 Then
vbwProfiler.vbwExecuteLine 3055
                            bFlag = True
vbwProfiler.vbwExecuteLine 3056
                            Exit For
                        End If
vbwProfiler.vbwExecuteLine 3057 'B
'vbwLine 3058:                    ElseIf element.material4 = "nonrigid" Then
                    ElseIf vbwProfiler.vbwExecuteLine(3058) Or element.material4 = "nonrigid" Then
vbwProfiler.vbwExecuteLine 3059
                         bReflex = True
vbwProfiler.vbwExecuteLine 3060
                        If element.dr4 > 100 Then
vbwProfiler.vbwExecuteLine 3061
                            bFlag = True
vbwProfiler.vbwExecuteLine 3062
                            Exit For
                        End If
vbwProfiler.vbwExecuteLine 3063 'B
'vbwLine 3064:                    ElseIf element.material5 = "nonrigid" Then
                    ElseIf vbwProfiler.vbwExecuteLine(3064) Or element.material5 = "nonrigid" Then
vbwProfiler.vbwExecuteLine 3065
                         bReflex = True
vbwProfiler.vbwExecuteLine 3066
                        If element.dr5 > 100 Then
vbwProfiler.vbwExecuteLine 3067
                            bFlag = True
vbwProfiler.vbwExecuteLine 3068
                            Exit For
                        End If
vbwProfiler.vbwExecuteLine 3069 'B
'vbwLine 3070:                    ElseIf element.material6 = "nonrigid" Then
                    ElseIf vbwProfiler.vbwExecuteLine(3070) Or element.material6 = "nonrigid" Then
vbwProfiler.vbwExecuteLine 3071
                         bReflex = True
vbwProfiler.vbwExecuteLine 3072
                        If element.dr6 > 100 Then
vbwProfiler.vbwExecuteLine 3073
                            bFlag = True
vbwProfiler.vbwExecuteLine 3074
                            Exit For
                        End If
vbwProfiler.vbwExecuteLine 3075 'B
                    End If
vbwProfiler.vbwExecuteLine 3076 'B

            End Select
vbwProfiler.vbwExecuteLine 3077 'B
        End If
vbwProfiler.vbwExecuteLine 3078 'B
vbwProfiler.vbwExecuteLine 3079
    Next

vbwProfiler.vbwExecuteLine 3080
    If bFlag Then

vbwProfiler.vbwExecuteLine 3081
        sText = "WARNING: You have set DR greater than 100 for an armor component which uses 'nonrigid' material.  Some GURPS users feel nonrigid armor DR cannot exceed 100."
    End If
vbwProfiler.vbwExecuteLine 3082 'B

vbwProfiler.vbwExecuteLine 3083
    If bReflex Then
vbwProfiler.vbwExecuteLine 3084
        If sText <> "" Then
vbwProfiler.vbwExecuteLine 3085
            sText = sText & vbNewLine & "NOTE: If you intend for your nonrigid armor to be 'reflex' armor, the DR should be limited to 5 x TL (VE page 22)"
        Else
vbwProfiler.vbwExecuteLine 3086 'B


vbwProfiler.vbwExecuteLine 3087
            sText = "NOTE: If you intend for your nonrigid armor to be 'reflex' armor, the DR should be limited to 5 x TL (VE page 22)"
        End If
vbwProfiler.vbwExecuteLine 3088 'B
    End If
vbwProfiler.vbwExecuteLine 3089 'B

vbwProfiler.vbwExecuteLine 3090
    MaxNonRigidArmorDRCheck = sText

vbwProfiler.vbwProcOut 115
vbwProfiler.vbwExecuteLine 3091
End Function

Function ArmorCheck() As String
vbwProfiler.vbwProcIn 116

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

vbwProfiler.vbwExecuteLine 3092
    With m_oCurrentVeh.surface
vbwProfiler.vbwExecuteLine 3093
        bSubmersible = .Submersible
vbwProfiler.vbwExecuteLine 3094
        bFloatationHull = .FloatationHull
vbwProfiler.vbwExecuteLine 3095
        If .StreamLining <> "none" Then
vbwProfiler.vbwExecuteLine 3096
            bStreamlining = True
        End If
vbwProfiler.vbwExecuteLine 3097 'B
vbwProfiler.vbwExecuteLine 3098
    End With

vbwProfiler.vbwExecuteLine 3099
    With m_oCurrentVeh.Components(BODY_KEY)
        ' check for slope on body
vbwProfiler.vbwExecuteLine 3100
        If .SlopeR <> "none" Then
vbwProfiler.vbwExecuteLine 3101
            bSlope = True
'vbwLine 3102:        ElseIf .slopel <> "none" Then
        ElseIf vbwProfiler.vbwExecuteLine(3102) Or .slopel <> "none" Then
vbwProfiler.vbwExecuteLine 3103
            bSlope = True
'vbwLine 3104:        ElseIf .slopef <> "none" Then
        ElseIf vbwProfiler.vbwExecuteLine(3104) Or .slopef <> "none" Then
vbwProfiler.vbwExecuteLine 3105
            bSlope = True
'vbwLine 3106:        ElseIf .slopeb <> "none" Then
        ElseIf vbwProfiler.vbwExecuteLine(3106) Or .slopeb <> "none" Then
vbwProfiler.vbwExecuteLine 3107
            bSlope = True
        End If
vbwProfiler.vbwExecuteLine 3108 'B
vbwProfiler.vbwExecuteLine 3109
    End With

    ' get array of all subassembly keys
vbwProfiler.vbwExecuteLine 3110
    arrSubs = m_oCurrentVeh.keymanager.GetCurrentSubAssembliesKeys


' check if any of these subs is a rotor or has slope
vbwProfiler.vbwExecuteLine 3111
If arrSubs(1) <> "" Then
vbwProfiler.vbwExecuteLine 3112
    For j = 1 To UBound(arrSubs)
vbwProfiler.vbwExecuteLine 3113
        If TypeOf m_oCurrentVeh.Components(arrSubs(j)) Is clsrotor Then
vbwProfiler.vbwExecuteLine 3114
            bRotors = True
        Else
vbwProfiler.vbwExecuteLine 3115 'B
            ' check to see if the vehilce has slope anywhere
vbwProfiler.vbwExecuteLine 3116
            If Not bSlope Then
vbwProfiler.vbwExecuteLine 3117
                If (TypeOf m_oCurrentVeh.Components(arrSubs(j)) Is clsBody) Or _
                    (TypeOf m_oCurrentVeh.Components(arrSubs(j)) Is clsSuperStructure) Or _
                    (TypeOf m_oCurrentVeh.Components(arrSubs(j)) Is clsTurret) Or _
                    (TypeOf m_oCurrentVeh.Components(arrSubs(j)) Is clsPopTurret) Then

vbwProfiler.vbwExecuteLine 3118
                    With m_oCurrentVeh.Components(arrSubs(j))
vbwProfiler.vbwExecuteLine 3119
                        If .SlopeR <> "none" Then
vbwProfiler.vbwExecuteLine 3120
                            bSlope = True
'vbwLine 3121:                        ElseIf .slopel <> "none" Then
                        ElseIf vbwProfiler.vbwExecuteLine(3121) Or .slopel <> "none" Then
vbwProfiler.vbwExecuteLine 3122
                            bSlope = True
'vbwLine 3123:                        ElseIf .slopef <> "none" Then
                        ElseIf vbwProfiler.vbwExecuteLine(3123) Or .slopef <> "none" Then
vbwProfiler.vbwExecuteLine 3124
                            bSlope = True
'vbwLine 3125:                        ElseIf .slopeb <> "none" Then
                        ElseIf vbwProfiler.vbwExecuteLine(3125) Or .slopeb <> "none" Then
vbwProfiler.vbwExecuteLine 3126
                            bSlope = True
                        End If
vbwProfiler.vbwExecuteLine 3127 'B
vbwProfiler.vbwExecuteLine 3128
                    End With
                End If
vbwProfiler.vbwExecuteLine 3129 'B
            End If
vbwProfiler.vbwExecuteLine 3130 'B

        ' bail out of For Loop early if possible
vbwProfiler.vbwExecuteLine 3131
        If bSlope And bRotors Then
vbwProfiler.vbwExecuteLine 3132
             Exit For
        End If
vbwProfiler.vbwExecuteLine 3133 'B
        End If
vbwProfiler.vbwExecuteLine 3134 'B
vbwProfiler.vbwExecuteLine 3135
    Next
End If
vbwProfiler.vbwExecuteLine 3136 'B

'now loop thru all subassemblies and make sure they have armor
vbwProfiler.vbwExecuteLine 3137
If arrSubs(1) <> "" Then
vbwProfiler.vbwExecuteLine 3138
    For j = 1 To UBound(arrSubs)
vbwProfiler.vbwExecuteLine 3139
        sKey = m_oCurrentVeh.Components(arrSubs(j)).Key

vbwProfiler.vbwExecuteLine 3140
        For Each element In m_oCurrentVeh.Components

vbwProfiler.vbwExecuteLine 3141
            If TypeOf element Is clsArmor Then
vbwProfiler.vbwExecuteLine 3142
                If element.Datatype = ArmorOverall Then
vbwProfiler.vbwProcOut 116
vbwProfiler.vbwExecuteLine 3143
                    Exit Function

'vbwLine 3144:                ElseIf element.LogicalParent = sKey Then
                ElseIf vbwProfiler.vbwExecuteLine(3144) Or element.LogicalParent = sKey Then
vbwProfiler.vbwExecuteLine 3145
                    bCurrentSubOK = True ' this subassembly does have armor. Keep checking though to see if we find an "overall armor"


                End If
vbwProfiler.vbwExecuteLine 3146 'B
            End If
vbwProfiler.vbwExecuteLine 3147 'B
vbwProfiler.vbwExecuteLine 3148
        Next

vbwProfiler.vbwExecuteLine 3149
        If Not bCurrentSubOK Then
vbwProfiler.vbwExecuteLine 3150
            Exit For ' bail out.  We do not have armor on all subs
        Else
vbwProfiler.vbwExecuteLine 3151 'B
vbwProfiler.vbwExecuteLine 3152
            bCurrentSubOK = False ' reset and test next sub
        End If
vbwProfiler.vbwExecuteLine 3153 'B
vbwProfiler.vbwExecuteLine 3154
    Next
End If
vbwProfiler.vbwExecuteLine 3155 'B

' if we havented exited this function, then we do not have necessary armor
' so we generate our message
vbwProfiler.vbwExecuteLine 3156
If bRotors Then
vbwProfiler.vbwExecuteLine 3157
     sText = " rotors,"
End If
vbwProfiler.vbwExecuteLine 3158 'B
vbwProfiler.vbwExecuteLine 3159
If bSubmersible Then
vbwProfiler.vbwExecuteLine 3160
     sText = sText & " submersible hull,"
End If
vbwProfiler.vbwExecuteLine 3161 'B
vbwProfiler.vbwExecuteLine 3162
If bFloatationHull Then 'todo: floationhull should be determined if the floatationrating is greater than 0!  This depends
vbwProfiler.vbwExecuteLine 3163
     sText = sText & " floatation hull,"
End If
vbwProfiler.vbwExecuteLine 3164 'B
vbwProfiler.vbwExecuteLine 3165
If bStreamlining Then
vbwProfiler.vbwExecuteLine 3166
     sText = sText & " streamlining,"
End If
vbwProfiler.vbwExecuteLine 3167 'B
vbwProfiler.vbwExecuteLine 3168
If bSlope Then
vbwProfiler.vbwExecuteLine 3169
     sText = sText & " slope,"
End If
vbwProfiler.vbwExecuteLine 3170 'B

vbwProfiler.vbwExecuteLine 3171
If sText = "" Then
vbwProfiler.vbwProcOut 116
vbwProfiler.vbwExecuteLine 3172
     Exit Function
End If
vbwProfiler.vbwExecuteLine 3173 'B

vbwProfiler.vbwExecuteLine 3174
sText = "Your vehicle has " & sText & " this requires all subassemblies to be armored or that 'overall armor' is used. Please add armor."

vbwProfiler.vbwExecuteLine 3175
ArmorCheck = sText

vbwProfiler.vbwProcOut 116
vbwProfiler.vbwExecuteLine 3176
End Function


Function DuplicateControlsCheck() As String
vbwProfiler.vbwProcIn 117

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

vbwProfiler.vbwExecuteLine 3177
    On Error Resume Next

vbwProfiler.vbwExecuteLine 3178
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 3179
        If TypeOf element Is clsManeuverControl Then
vbwProfiler.vbwExecuteLine 3180
            Select Case element.Datatype

'vbwLine 3181:                Case ElectronicDivingControl
                Case IIf(vbwProfiler.vbwExecuteLine(3181), VBWPROFILER_EMPTY, _
        ElectronicDivingControl)
vbwProfiler.vbwExecuteLine 3182
                    If element.duplicate Then

vbwProfiler.vbwExecuteLine 3183
                        bElectricDivingDup = True
                    Else
vbwProfiler.vbwExecuteLine 3184 'B
vbwProfiler.vbwExecuteLine 3185
                        bElectricDiving = True
                    End If
vbwProfiler.vbwExecuteLine 3186 'B

'vbwLine 3187:                Case ElectronicManeuverControl
                Case IIf(vbwProfiler.vbwExecuteLine(3187), VBWPROFILER_EMPTY, _
        ElectronicManeuverControl)
vbwProfiler.vbwExecuteLine 3188
                    If element.duplicate Then
vbwProfiler.vbwExecuteLine 3189
                        bElectricDup = True
                    Else
vbwProfiler.vbwExecuteLine 3190 'B
vbwProfiler.vbwExecuteLine 3191
                        bElectric = True
                    End If
vbwProfiler.vbwExecuteLine 3192 'B

'vbwLine 3193:                Case ComputerizedDivingControl
                Case IIf(vbwProfiler.vbwExecuteLine(3193), VBWPROFILER_EMPTY, _
        ComputerizedDivingControl)
vbwProfiler.vbwExecuteLine 3194
                    If element.duplicate Then
vbwProfiler.vbwExecuteLine 3195
                        bComputerizedDivingDup = True
                    Else
vbwProfiler.vbwExecuteLine 3196 'B
vbwProfiler.vbwExecuteLine 3197
                        bComputerizedDiving = True
                    End If
vbwProfiler.vbwExecuteLine 3198 'B

'vbwLine 3199:                Case ComputerizedManeuverControl
                Case IIf(vbwProfiler.vbwExecuteLine(3199), VBWPROFILER_EMPTY, _
        ComputerizedManeuverControl)
vbwProfiler.vbwExecuteLine 3200
                    If element.duplicate Then
vbwProfiler.vbwExecuteLine 3201
                        bComputerizedDup = True
                    Else
vbwProfiler.vbwExecuteLine 3202 'B
vbwProfiler.vbwExecuteLine 3203
                        bComputerized = True
                    End If
vbwProfiler.vbwExecuteLine 3204 'B

'vbwLine 3205:                Case MechanicalManeuverControl
                Case IIf(vbwProfiler.vbwExecuteLine(3205), VBWPROFILER_EMPTY, _
        MechanicalManeuverControl)
vbwProfiler.vbwExecuteLine 3206
                    If element.duplicate Then
vbwProfiler.vbwExecuteLine 3207
                        bMechanicalDup = True
                    Else
vbwProfiler.vbwExecuteLine 3208 'B
vbwProfiler.vbwExecuteLine 3209
                        bMechanical = True
                    End If
vbwProfiler.vbwExecuteLine 3210 'B

'vbwLine 3211:                Case MechanicalDivingControl
                Case IIf(vbwProfiler.vbwExecuteLine(3211), VBWPROFILER_EMPTY, _
        MechanicalDivingControl)
vbwProfiler.vbwExecuteLine 3212
                    If element.duplicate Then
vbwProfiler.vbwExecuteLine 3213
                        bMechanicalDivingDup = True
                    Else
vbwProfiler.vbwExecuteLine 3214 'B
vbwProfiler.vbwExecuteLine 3215
                        bMechanicalDiving = True
                    End If
vbwProfiler.vbwExecuteLine 3216 'B
            End Select
vbwProfiler.vbwExecuteLine 3217 'B
        End If
vbwProfiler.vbwExecuteLine 3218 'B
vbwProfiler.vbwExecuteLine 3219
    Next


vbwProfiler.vbwExecuteLine 3220
    If (bMechanicalDivingDup) And (bMechanicalDiving = False) Then
vbwProfiler.vbwExecuteLine 3221
        sRet = "You have a duplicate Mechanical Diving control but no primary (non duplicate) set installed."
    End If
vbwProfiler.vbwExecuteLine 3222 'B

vbwProfiler.vbwExecuteLine 3223
    If (bMechanicalDup) And (bMechanical = False) Then
vbwProfiler.vbwExecuteLine 3224
        If sRet <> "" Then
vbwProfiler.vbwExecuteLine 3225
             sRet = sRet & vbNewLine
        End If
vbwProfiler.vbwExecuteLine 3226 'B
vbwProfiler.vbwExecuteLine 3227
        sRet = sRet & "You have a duplicate Mechanical maneuver control but no primary (non duplicate) set installed."
    End If
vbwProfiler.vbwExecuteLine 3228 'B

vbwProfiler.vbwExecuteLine 3229
    If (bComputerizedDup) And (bComputerized = False) Then
vbwProfiler.vbwExecuteLine 3230
        If sRet <> "" Then
vbwProfiler.vbwExecuteLine 3231
             sRet = sRet & vbNewLine
        End If
vbwProfiler.vbwExecuteLine 3232 'B
vbwProfiler.vbwExecuteLine 3233
        sRet = sRet & "You have a duplicate Computerized maneuver control but no primary (non duplicate) set installed."
    End If
vbwProfiler.vbwExecuteLine 3234 'B

vbwProfiler.vbwExecuteLine 3235
    If (bComputerizedDivingDup) And (bComputerizedDiving = False) Then
vbwProfiler.vbwExecuteLine 3236
        If sRet <> "" Then
vbwProfiler.vbwExecuteLine 3237
             sRet = sRet & vbNewLine
        End If
vbwProfiler.vbwExecuteLine 3238 'B
vbwProfiler.vbwExecuteLine 3239
        sRet = sRet & "You have a duplicate Computerized diving control but no primary (non duplicate) set installed."
    End If
vbwProfiler.vbwExecuteLine 3240 'B

vbwProfiler.vbwExecuteLine 3241
    If (bElectricDup) And (bElectric = False) Then
vbwProfiler.vbwExecuteLine 3242
        If sRet <> "" Then
vbwProfiler.vbwExecuteLine 3243
             sRet = sRet & vbNewLine
        End If
vbwProfiler.vbwExecuteLine 3244 'B
vbwProfiler.vbwExecuteLine 3245
        sRet = sRet & "You have a duplicate Electronic maneuver control but no primary (non duplicate) set installed."
    End If
vbwProfiler.vbwExecuteLine 3246 'B

vbwProfiler.vbwExecuteLine 3247
    If (bElectricDivingDup) And (bElectricDiving = False) Then
vbwProfiler.vbwExecuteLine 3248
        If sRet <> "" Then
vbwProfiler.vbwExecuteLine 3249
             sRet = sRet & vbNewLine
        End If
vbwProfiler.vbwExecuteLine 3250 'B
vbwProfiler.vbwExecuteLine 3251
        sRet = sRet & "You have a duplicate Electronic diving control but no primary (non duplicate) set installed."
    End If
vbwProfiler.vbwExecuteLine 3252 'B

vbwProfiler.vbwExecuteLine 3253
   DuplicateControlsCheck = sRet
vbwProfiler.vbwProcOut 117
vbwProfiler.vbwExecuteLine 3254
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
vbwProfiler.vbwProcIn 118
vbwProfiler.vbwProcOut 118
vbwProfiler.vbwExecuteLine 3255
End Function

Function LegCheck() As String
vbwProfiler.vbwProcIn 119
    Dim MinLegVolume As Single ' each leg must be (0.04 x body volume / number of legs) in volume
    Dim arrLegVolumes() As Single ' array to hold the volumes for all the legs
    Dim legarray() As String 'holds the keys for all legs in the vehicle
    Dim i, NumLegs, j As Integer ' counter
    Dim retval As String
    Dim retval2 As String

vbwProfiler.vbwExecuteLine 3256
    On Error GoTo errorhandler
    ' Find how many legs are on the vehicle
vbwProfiler.vbwExecuteLine 3257
    legarray = m_oCurrentVeh.Components(BODY_KEY).GetCurrentLegKeys
vbwProfiler.vbwExecuteLine 3258
    NumLegs = UBound(legarray)
vbwProfiler.vbwExecuteLine 3259
    If legarray(1) <> "" Then
vbwProfiler.vbwExecuteLine 3260
        ReDim arrLegVolumes(NumLegs) 'set the dimension of the array to the number of legs
        ' Check to see if Leg minimum volume is attained
vbwProfiler.vbwExecuteLine 3261
        MinLegVolume = (m_oCurrentVeh.Components(BODY_KEY).Volume * 0.04 / NumLegs)
vbwProfiler.vbwExecuteLine 3262
        MinLegVolume = Round(MinLegVolume, 2)

        ' set the volume of the first leg which will be used to compare volumes of others
vbwProfiler.vbwExecuteLine 3263
        arrLegVolumes(1) = m_oCurrentVeh.Components(legarray(1)).Volume

        'check to see that each leg is the bigger than min volume
vbwProfiler.vbwExecuteLine 3264
        For i = 1 To NumLegs
vbwProfiler.vbwExecuteLine 3265
            If m_oCurrentVeh.Components(legarray(i)).Volume < MinLegVolume Then
vbwProfiler.vbwExecuteLine 3266
                retval = "Each leg must be greater than " & MinLegVolume & " cf. of volume.  Add more legs to balance the load or add empty space to the legs to increase their volume."
            End If
vbwProfiler.vbwExecuteLine 3267 'B
vbwProfiler.vbwExecuteLine 3268
        Next
        ' Check to see that each leg is also the same volume
vbwProfiler.vbwExecuteLine 3269
        For j = 2 To NumLegs
vbwProfiler.vbwExecuteLine 3270
            If arrLegVolumes(1) <> m_oCurrentVeh.Components(legarray(j)).Volume Then
vbwProfiler.vbwExecuteLine 3271
                retval2 = "Each leg must be the same volume as the other legs.  Check volumes of each leg and add empty space if necessary"
            End If
vbwProfiler.vbwExecuteLine 3272 'B
vbwProfiler.vbwExecuteLine 3273
        Next
    End If
vbwProfiler.vbwExecuteLine 3274 'B

vbwProfiler.vbwExecuteLine 3275
    If retval = "" Then
vbwProfiler.vbwExecuteLine 3276
        retval = retval2
    Else
vbwProfiler.vbwExecuteLine 3277 'B
vbwProfiler.vbwExecuteLine 3278
        retval = retval & " " & retval2
    End If
vbwProfiler.vbwExecuteLine 3279 'B
vbwProfiler.vbwExecuteLine 3280
    LegCheck = retval
vbwProfiler.vbwProcOut 119
vbwProfiler.vbwExecuteLine 3281
    Exit Function

errorhandler:

vbwProfiler.vbwProcOut 119
vbwProfiler.vbwExecuteLine 3282
End Function

Function HardpointCapacityCheck() As String
vbwProfiler.vbwProcIn 120
    Dim retval As String
    Dim element As Object
vbwProfiler.vbwExecuteLine 3283
    On Error GoTo errorhandler
    Dim sParentKey As String

vbwProfiler.vbwExecuteLine 3284
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 3285
        sParentKey = element.LogicalParent
vbwProfiler.vbwExecuteLine 3286
        If sParentKey <> "" Then
vbwProfiler.vbwExecuteLine 3287
            If (TypeOf m_oCurrentVeh.Components(sParentKey) Is clsHardPoint) Then
vbwProfiler.vbwExecuteLine 3288
                If element.Weight > m_oCurrentVeh.Components(sParentKey).loadcapacity Then
vbwProfiler.vbwExecuteLine 3289
                    retval = retval & element.Description & " weight exceeds the maximum capacity of the hardpoint or weaponbay its attached to. "
                End If
vbwProfiler.vbwExecuteLine 3290 'B
            End If
vbwProfiler.vbwExecuteLine 3291 'B
        End If
vbwProfiler.vbwExecuteLine 3292 'B
vbwProfiler.vbwExecuteLine 3293
    Next
vbwProfiler.vbwExecuteLine 3294
    HardpointCapacityCheck = retval
vbwProfiler.vbwProcOut 120
vbwProfiler.vbwExecuteLine 3295
    Exit Function

errorhandler:



vbwProfiler.vbwProcOut 120
vbwProfiler.vbwExecuteLine 3296
End Function


Function SolarCellArrayCheck() As String
vbwProfiler.vbwProcIn 121
    Dim retval As String
    Dim element As Object
vbwProfiler.vbwExecuteLine 3297
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
vbwProfiler.vbwExecuteLine 3298
    With m_oCurrentVeh.surface
vbwProfiler.vbwExecuteLine 3299
        TL = m_oCurrentVeh.Components(BODY_KEY).TL
vbwProfiler.vbwExecuteLine 3300
        If .infraredcloaking <> "none" Then
vbwProfiler.vbwExecuteLine 3301
             bInfrared = True
        End If
vbwProfiler.vbwExecuteLine 3302 'B
vbwProfiler.vbwExecuteLine 3303
        If .Stealth <> "none" Then
vbwProfiler.vbwExecuteLine 3304
             bStealth = True
        End If
vbwProfiler.vbwExecuteLine 3305 'B
vbwProfiler.vbwExecuteLine 3306
        If .Chameleon <> "none" Then
vbwProfiler.vbwExecuteLine 3307
             bChameleon = True
        End If
vbwProfiler.vbwExecuteLine 3308 'B
vbwProfiler.vbwExecuteLine 3309
        If Not .LiquidCrystal Then
vbwProfiler.vbwExecuteLine 3310
             bLiquidCrystal = True
        End If
vbwProfiler.vbwExecuteLine 3311 'B
vbwProfiler.vbwExecuteLine 3312
    End With

    ' check that only TL10+ can have solar cells on gasbags
vbwProfiler.vbwExecuteLine 3313
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 3314
        If TypeOf element Is clsSolarCellArray Then
vbwProfiler.vbwExecuteLine 3315
            bCells = True
vbwProfiler.vbwExecuteLine 3316
            sParentKey = element.LogicalParent
vbwProfiler.vbwExecuteLine 3317
            If sParentKey <> "" Then
vbwProfiler.vbwExecuteLine 3318
                If (TypeOf m_oCurrentVeh.Components(sParentKey) Is clsGasbag) Then
vbwProfiler.vbwExecuteLine 3319
                    If TL < 10 Then
vbwProfiler.vbwExecuteLine 3320
                        retval = retval & element.Description & " should not be placed on gasbags prior to TL 10. "
vbwProfiler.vbwExecuteLine 3321
                        bNotOnPanels = True
vbwProfiler.vbwExecuteLine 3322
                        Exit For
                    End If
vbwProfiler.vbwExecuteLine 3323 'B
'vbwLine 3324:                ElseIf Not TypeOf m_oCurrentVeh.Components(sParentKey) Is clsSolarPanel Then
                ElseIf vbwProfiler.vbwExecuteLine(3324) Or Not TypeOf m_oCurrentVeh.Components(sParentKey) Is clsSolarPanel Then
vbwProfiler.vbwExecuteLine 3325
                    bNotOnPanels = True
                End If
vbwProfiler.vbwExecuteLine 3326 'B
            End If
vbwProfiler.vbwExecuteLine 3327 'B
        End If
vbwProfiler.vbwExecuteLine 3328 'B
vbwProfiler.vbwExecuteLine 3329
    Next

vbwProfiler.vbwExecuteLine 3330
    If (bCells) And (TL < 11) And (bNotOnPanels) Then
vbwProfiler.vbwExecuteLine 3331
        If (bInfrared) Or (bStealth) Or (bChameleon) Or (bLiquidCrystal) Then
vbwProfiler.vbwExecuteLine 3332
            If retval <> "" Then
vbwProfiler.vbwExecuteLine 3333
                retval = retval & vbNewLine
            End If
vbwProfiler.vbwExecuteLine 3334 'B

vbwProfiler.vbwExecuteLine 3335
            retval = retval & "Vehicles with stealth, infrared cloaking, chameleon systems or liquid crystal skins should not have solar cells on them prior to TL11 unless they are on solar panels."

        End If
vbwProfiler.vbwExecuteLine 3336 'B
    End If
vbwProfiler.vbwExecuteLine 3337 'B

vbwProfiler.vbwExecuteLine 3338
    SolarCellArrayCheck = retval
vbwProfiler.vbwProcOut 121
vbwProfiler.vbwExecuteLine 3339
    Exit Function

errorhandler:



vbwProfiler.vbwProcOut 121
vbwProfiler.vbwExecuteLine 3340
End Function

Function TopSpeedRestrictions() As String
vbwProfiler.vbwProcIn 122
    Dim pType As String
    Dim element As Object
    Dim sRet As String

vbwProfiler.vbwExecuteLine 3341
    For Each element In m_oCurrentVeh.PerformanceProfiles

vbwProfiler.vbwExecuteLine 3342
            Select Case element.Datatype
'vbwLine 3343:                Case PERFORMANCEHOVER, PERFORMANCEAIR, PERFORMANCEMAGLEV
                Case IIf(vbwProfiler.vbwExecuteLine(3343), VBWPROFILER_EMPTY, _
        PERFORMANCEHOVER), PERFORMANCEAIR, PERFORMANCEMAGLEV
vbwProfiler.vbwExecuteLine 3344
                    If element.DesignCheckString <> "" Then

vbwProfiler.vbwExecuteLine 3345
                        sRet = sRet & "FYI: In your performance profile named '" & element.Key & "' " & element.DesignCheckString
                    End If
vbwProfiler.vbwExecuteLine 3346 'B
            End Select
vbwProfiler.vbwExecuteLine 3347 'B


vbwProfiler.vbwExecuteLine 3348
    Next

vbwProfiler.vbwExecuteLine 3349
    TopSpeedRestrictions = sRet

vbwProfiler.vbwProcOut 122
vbwProfiler.vbwExecuteLine 3350
End Function
Function SoftwareComplexityCheck() As String
    '//check to make sure that software is not of greater complexity
    '//then the computer its added to
vbwProfiler.vbwProcIn 123
    Dim element As Object
    Dim retval As String
    Dim sParent As String

vbwProfiler.vbwExecuteLine 3351
    On Error GoTo errorhandler

vbwProfiler.vbwExecuteLine 3352
    For Each element In m_oCurrentVeh.Components
vbwProfiler.vbwExecuteLine 3353
        sParent = element.LogicalParent
vbwProfiler.vbwExecuteLine 3354
        If sParent <> "" Then
vbwProfiler.vbwExecuteLine 3355
            If TypeOf m_oCurrentVeh.Components(sParent) Is clsComputer Then
vbwProfiler.vbwExecuteLine 3356
                If element.complexity > m_oCurrentVeh.Components(sParent).complexity Then
vbwProfiler.vbwExecuteLine 3357
                    retval = retval & element.Description & " complexity exceeds the complexity rating of the " & element.Description & " its attached to. "
                End If
vbwProfiler.vbwExecuteLine 3358 'B
            End If
vbwProfiler.vbwExecuteLine 3359 'B
        End If
vbwProfiler.vbwExecuteLine 3360 'B
vbwProfiler.vbwExecuteLine 3361
    Next
vbwProfiler.vbwExecuteLine 3362
    SoftwareComplexityCheck = retval
vbwProfiler.vbwProcOut 123
vbwProfiler.vbwExecuteLine 3363
    Exit Function

errorhandler:

vbwProfiler.vbwProcOut 123
vbwProfiler.vbwExecuteLine 3364
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

