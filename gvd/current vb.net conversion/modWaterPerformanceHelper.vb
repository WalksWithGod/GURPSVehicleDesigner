Option Strict Off
Option Explicit On
Module modWaterPerformanceHelper
	
	Public Function CalcWaterAcceleration(ByVal Thrust As Double, ByVal LoadedWeight As Double) As Single
		Dim TempAcceleration As Single
		
		If LoadedWeight = 0 Then Exit Function
		
		TempAcceleration = (Thrust / LoadedWeight) * 20
		If TempAcceleration < 1 Then
			' round to 1 decimal place
			TempAcceleration = System.Math.Round(TempAcceleration, 1)
			' if less than .1  make sure its at least .1
			'UPGRADE_WARNING: Couldn't resolve default property of object Maximum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			TempAcceleration = modPerformance.Maximum(TempAcceleration, 0.1)
		ElseIf TempAcceleration < 5 Then 
			' round to nearest 1mph
			TempAcceleration = System.Math.Round(TempAcceleration, 0)
		Else
			' round to nearest 5mph
			TempAcceleration = System.Math.Round(TempAcceleration * 5, 0) \ 5
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
	
	
	Public Function GetHl() As Short
		Dim sLines As String
		Dim Hl As Short
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
				Debug.Print("modWaterPerformanceHelper:GetHL() -- ERROR.  Invalid Case")
		End Select
		
		GetHl = Hl
	End Function
	
	Sub CalcWaterMRandSR(ByRef SR As Single, ByRef MR As Single, ByRef KeyChain As Object, ByVal bResponsive As Boolean)
		Dim Catamaran As Object
		Dim Trimaran As Object
		Dim Category As Short
		Dim TempMR As Single
		Dim TempSR As Single
		Dim TempType As udtWaterStability
		Dim TempTech As Short
		Dim MRBonus As Single
		Dim SRBonus As Short
		Dim dType As String
		Dim sKey As String
		Dim sLines As String
		
		Dim i As Short
		
		'UPGRADE_WARNING: Couldn't resolve default property of object KeyChain(1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If KeyChain(1) = "" Then Exit Sub 'dont perform routine if there are no propulsion systems in the keychain
		
		MRBonus = 0 'init the bonus value
		SRBonus = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With Veh.Components(BODY_KEY)
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			TempTech = .TL
			If TempTech > 11 Then TempTech = 11 '11 is the highest
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .Volume <= 100 Then
				Category = 1
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf .Volume <= 1000 Then 
				Category = 2
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf .Volume <= 10000 Then 
				Category = 3
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf .Volume <= 100000 Then 
				Category = 4
				'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf .Volume <= 1000000 Then 
				Category = 5
			Else : Category = 6
			End If
		End With
		'get the initial values
		'UPGRADE_WARNING: Couldn't resolve default property of object TempType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
			'UPGRADE_WARNING: Couldn't resolve default property of object KeyChain(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sKey = KeyChain(i)
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dType = Veh.Components(sKey).Datatype
			
			'check for Flexibody drivetrain modifier
			If CDbl(dType) = FlexibodyDrivetrain Then
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
		'UPGRADE_WARNING: Couldn't resolve default property of object TempType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		TempType = WaterStabMatrix(TempTech, Category)
		MR = TempType.MR + MRBonus
		
		'make the SR adjustments
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Veh.Options.RollStabilizers Then SRBonus = SRBonus + 1
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Veh.surface.CataTrimaran = Catamaran Then
			SRBonus = SRBonus + 2
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		ElseIf Veh.surface.CataTrimaran = Trimaran Then 
			SRBonus = SRBonus + 2
		End If
		
		
		'set the final value for stability
		TempSR = TempSR + SRBonus
		If TempSR < 1 Then TempSR = 1
		SR = TempSR
	End Sub
End Module