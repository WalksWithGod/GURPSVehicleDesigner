Attribute VB_Name = "modWaterPerformanceHelper"
Option Explicit

Public Function CalcWaterAcceleration(ByVal Thrust As Double, ByVal LoadedWeight As Double) As Single
    Dim TempAcceleration As Single
    
    If LoadedWeight = 0 Then Exit Function
    
    TempAcceleration = (Thrust / LoadedWeight) * 20
    If TempAcceleration < 1 Then
        ' round to 1 decimal place
        TempAcceleration = Round(TempAcceleration, 1)
        ' if less than .1  make sure its at least .1
        TempAcceleration = Maximum(TempAcceleration, 0.1)
    ElseIf TempAcceleration < 5 Then
        ' round to nearest 1mph
        TempAcceleration = Round(TempAcceleration, 0)
    Else
        ' round to nearest 5mph
        TempAcceleration = Round(TempAcceleration * 5, 0) \ 5
    End If
    CalcWaterAcceleration = TempAcceleration
End Function



Sub CalcWaterDeceleration(ByVal MR As Single, ByVal Accel As Single, ByRef Decel As Single, ByRef PoweredDecel As Single)
    Dim Temp As Single

    Temp = 100 * (MR / GetHl)
    If Temp > 10 Then Temp = 10
    Decel = Temp ' this is for unpowered/drifting deceleration
    PoweredDecel = (Accel / 2) + Temp ' this is for powered deceleration
End Sub


Public Function GetHl() As Integer
    Dim sLines As String
    Dim Hl As Integer
    
    sLines = Veh.surface.HydrodynamicLines
    
    Select Case sLines
        Case "very fine"
            Hl = 20
        Case "fine"
            Hl = 15
        Case "average"
            Hl = 10
        Case "mediocre"
            Hl = 5
        Case "none"
            Hl = 1
        Case "submarine"
            Hl = 5
        Case Else
            Debug.Print "modWaterPerformanceHelper:GetHL() -- ERROR.  Invalid Case"
    End Select

    GetHl = Hl
End Function

Sub CalcWaterMRandSR(ByRef SR As Single, ByRef MR As Single, ByRef KeyChain As Variant, ByVal bResponsive As Boolean)
    Dim Category As Integer
    Dim TempMR As Single
    Dim TempSR As Single
    Dim TempType As udtWaterStability
    Dim TempTech As Integer
    Dim MRBonus As Single
    Dim SRBonus As Integer
    Dim dType As String
    Dim sKey As String
    Dim sLines As String
    
    Dim i As Integer
    
    If KeyChain(1) = "" Then Exit Sub 'dont perform routine if there are no propulsion systems in the keychain
    
    MRBonus = 0 'init the bonus value
    SRBonus = 0
    With Veh.Components(BODY_KEY)
        TempTech = .TL
        If TempTech > 11 Then TempTech = 11 '11 is the highest
        If .Volume <= 100 Then
            Category = 1
        ElseIf .Volume <= 1000 Then Category = 2
        ElseIf .Volume <= 10000 Then Category = 3
        ElseIf .Volume <= 100000 Then Category = 4
        ElseIf .Volume <= 1000000 Then Category = 5
        Else: Category = 6
        End If
    End With
    'get the initial values
    TempType = WaterStabMatrix(TempTech, Category)
    TempSR = TempType.SR ' need to save this early since we will inadvertantly modify it
    
    'make the MR category adjustments
    If bResponsive Then
        If Category = 1 Then
            MRBonus = MRBonus + 0.25
        Else
            Category = Category - 1
        End If
    End If
    For i = 1 To UBound(KeyChain)
    sKey = KeyChain(i)
    dType = Veh.Components(sKey).Datatype
    
        'check for Flexibody drivetrain modifier
        If dType = FlexibodyDrivetrain Then
            If Category = 1 Then
                MRBonus = MRBonus + 0.25
            Else
                Category = Category - 1
            End If
        End If
    Next
     
    'add modifier for Electric or comptuerized controls
    If VehiclehasElectORCompcontrols Then
        If Category = 1 Then
            MRBonus = MRBonus + 0.25
        Else
            Category = Category - 1
        End If
        SRBonus = SRBonus + 1 ' one of the SR modifiers
    End If
        
    'set the final value for maneuverability
    TempType = WaterStabMatrix(TempTech, Category)
    MR = TempType.MR + MRBonus
    
    'make the SR adjustments
    
    If Veh.Options.RollStabilizers Then SRBonus = SRBonus + 1
    
    sLines = Veh.surface.HydrodynamicLines
    Select Case sLines
        Case "average"
            SRBonus = SRBonus - 1
        Case "fine"
            SRBonus = SRBonus - 2
        Case "very fine"
            SRBonus = SRBonus - 2
        Case "submarine"
            SRBonus = SRBonus - 2
    End Select
    
    If Veh.surface.CataTrimaran = Catamaran Then
        SRBonus = SRBonus + 2
    ElseIf Veh.surface.CataTrimaran = Trimaran Then
        SRBonus = SRBonus + 2
    End If

    
    'set the final value for stability
    TempSR = TempSR + SRBonus
    If TempSR < 1 Then TempSR = 1
    SR = TempSR
End Sub


