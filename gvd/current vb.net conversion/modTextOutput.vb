Option Strict Off
Option Explicit On
Module modTextOutput
	Private sBreak As String
	Private sLineBreak As String
	Private bSlimline As Boolean
	
	Private Const CREATED_WITH As String = "Created with GURPS Vehicle Designer 2.0"
	Private Const GVD_URL As String = "http://www.makosoft.com/gvd"
	
	
	Public Function createGURPSText(ByVal sType As String) As String
		Dim sOutput As String
		Dim sTemp As String
		Dim sTagline As String
		
#If DEBUG_MODE Then
		'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression DEBUG_MODE did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
		MsgBox "modTextOutput:createGURPSText() - Function not available in Debug Mode."
		Exit Function
#End If
		
		'jaw 2000.06.25
		'reformed to select case to allow for additional exports to be easily added
		Select Case sType
			Case "Text"
				sBreak = Chr(13) & Chr(10) & Chr(13) & Chr(10)
				sTagline = CREATED_WITH & Chr(13) & Chr(10) & GVD_URL
				sLineBreak = Chr(13) & Chr(10)
			Case "Text Slim"
				sBreak = Chr(13) & Chr(10) '+ Chr(13) + Chr(10)
				sTagline = CREATED_WITH & Chr(13) & Chr(10) & GVD_URL & Chr(13) & Chr(10)
				sLineBreak = Chr(13) & Chr(10)
				bSlimline = True
			Case "Class HTML"
				sBreak = "<BR> <BR>" & vbNewLine & vbNewLine
				sTagline = CREATED_WITH & "<BR>" & GVD_URL & vbCrLf & "</body></html>"
				sLineBreak = "<BR>"
			Case "New HTML"
				sBreak = "<BR> <BR>" & vbNewLine & vbNewLine
				sTagline = CREATED_WITH & "<BR>" & GVD_URL & vbCrLf & "</body></html>"
				sLineBreak = "<BR>"
			Case Else
				Exit Function
		End Select
		On Error Resume Next
		'get header, vehicle name, copyright info and description
		sOutput = GetHeaderOutput(sType)
		sOutput = sOutput & "Subassemblies and Body Features: " & GetSubassemblyOutput & GetBodyFeatures & sBreak
		sTemp = GetCustomComponentsOutput
		If sTemp <> "" Then sOutput = sOutput & "Custom Components: " & sTemp & sBreak
		sTemp = GetPropulsionOutput
		If sTemp <> "" Then sOutput = sOutput & "Propulsion: " & sTemp & sBreak
		sTemp = GetAerostaticLiftOutput
		If sTemp <> "" Then sOutput = sOutput & "Aerostatic Lift: " & sTemp & sBreak
		sTemp = GetWeaponryOutput
		If sTemp <> "" Then sOutput = sOutput & "Weaponry: " & sTemp & sBreak
		sTemp = GetWeaponLinksOutput
		If sTemp <> "" Then sOutput = sOutput & "Weapon Links: " & sTemp & sBreak
		sTemp = GetWeaponAccessoriesOutput
		If sTemp <> "" Then sOutput = sOutput & "Weapon Accessories: " & sTemp & sBreak
		sTemp = GetCommunicationsOutput
		If sTemp <> "" Then sOutput = sOutput & "Communications: " & sTemp & sBreak
		sTemp = GetSensorsOutput
		If sTemp <> "" Then sOutput = sOutput & "Sensors: " & sTemp & sBreak
		sTemp = GetAudioVisualOutput
		If sTemp <> "" Then sOutput = sOutput & "Audio/Visual: " & sTemp & sBreak
		sTemp = GetNavigationOutput
		If sTemp <> "" Then sOutput = sOutput & "Navigation: " & sTemp & sBreak
		sTemp = GetTargetingOutput
		If sTemp <> "" Then sOutput = sOutput & "Targeting: " & sTemp & sBreak
		sTemp = GetECMOutput
		If sTemp <> "" Then sOutput = sOutput & "ECM: " & sTemp & sBreak
		sTemp = GetComputersOutput
		If sTemp <> "" Then sOutput = sOutput & "Computers: " & sTemp & sBreak
		sTemp = GetSoftwareOutput
		If sTemp <> "" Then sOutput = sOutput & "Software: " & sTemp & sBreak
		sTemp = GetMiscellaneousOutput
		If sTemp <> "" Then sOutput = sOutput & "Miscellaneous: " & sTemp & sBreak
		sTemp = GetVehicleControlsOutput
		If sTemp <> "" Then sOutput = sOutput & "Vehicle Controls: " & sTemp & sBreak
		sTemp = GetNeuralInterfaceSystemOutput
		If sTemp <> "" Then sOutput = sOutput & "Neural Interfaces: " & sTemp & sBreak
		sTemp = GetCrewStationsOutput
		If sTemp <> "" Then sOutput = sOutput & "Crew Stations: " & sTemp & sBreak
		sTemp = GetOccupancyOutput
		If sTemp <> "" Then sOutput = sOutput & "Occupancy: " & sTemp & sBreak
		sTemp = GetAccomodationsOutput
		If sTemp <> "" Then sOutput = sOutput & "Accommodations: " & sTemp & sBreak
		sTemp = GetEnvironmentalSystemsOutput
		If sTemp <> "" Then sOutput = sOutput & "Environmental Systems: " & sTemp & sBreak
		sTemp = GetSafetySystemsOutput
		If sTemp <> "" Then sOutput = sOutput & "Safety Systems: " & sTemp & sBreak
		sTemp = GetPowerSystemsOutPut
		If sTemp <> "" Then sOutput = sOutput & "Power Systems: " & sTemp & sBreak
		sTemp = GetFuelOutput
		If sTemp <> "" Then sOutput = sOutput & "Fuel: " & sTemp & sBreak
		sTemp = GetSpaceOutput
		If sTemp <> "" Then sOutput = sOutput & sTemp & sBreak
		sTemp = GetSurfaceAreaOutput
		If sTemp <> "" Then sOutput = sOutput & "Surface Area: " & sTemp & sBreak
		sTemp = GetStructureOutput
		If sTemp <> "" Then sOutput = sOutput & "Structure: " & sTemp & sBreak
		sTemp = GetHitPointsOutput
		If sTemp <> "" Then sOutput = sOutput & "Hit Points: " & sTemp & sBreak
		sTemp = GetStructuralOptionsOutput
		If sTemp <> "" Then sOutput = sOutput & sTemp & sBreak
		sTemp = GetArmorOutput
		If sTemp <> "" Then sOutput = sOutput & sTemp & sBreak
		sTemp = GetSurfaceFeaturesOutput
		If sTemp <> "" Then sOutput = sOutput & sTemp & sBreak
		sTemp = GetDefensiveSurfaceFeatures
		If sTemp <> "" Then sOutput = sOutput & sTemp & sBreak
		sTemp = GetOtherSurfaceFeatures
		If sTemp <> "" Then sOutput = sOutput & sTemp & sBreak
		sTemp = GetTopDeckSurfaceFeatures
		If sTemp <> "" Then sOutput = sOutput & sTemp & sBreak
		sTemp = GetWeaponBaysAndHardpoints
		If sTemp <> "" Then sOutput = sOutput & sTemp & sBreak
		sTemp = GetVisionAndDetailsOutput
		If sTemp <> "" Then sOutput = sOutput & sTemp & sBreak
		sTemp = GetStatisticsOutput
		If sTemp <> "" Then sOutput = sOutput & "Statistics: " & sTemp & sBreak
		sTemp = GetPerformanceOutput
		If sTemp <> "" Then sOutput = sOutput & sTemp & sBreak
		sTemp = GetDetailedWeaponStats
		If sTemp <> "" Then sOutput = sOutput & sTemp & vbNewLine
		
		'//add our tag line
		sOutput = sOutput & sTagline
		'jaw 2000.06.25
		If bSlimline Then
			'UPGRADE_WARNING: Couldn't resolve default property of object RemoveParenthetical(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sOutput = RemoveParenthetical(sOutput)
		End If
		createGURPSText = sOutput
	End Function
	
	Private Function GetHeaderOutput(ByVal sType As String) As String
		Dim sOutput As String
		Dim sTemp As String
		'    With m_oCurrentVeh.Description
		'        sTemp = .Header
		'        If sTemp <> "" Then sOutput = sTemp + sBreak
		'        sTemp = .NickName
		'        If sTemp <> "" Then sOutput = sOutput + "Name: " + sTemp + sLineBreak
		'        sTemp = .ClassName
		'        If sTemp <> "" Then sOutput = sOutput + "Class: " & sTemp + sLineBreak
		'        sTemp = .category
		'        If sTemp <> "" Then
		'            If .subcategory <> "" Then
		'                sOutput = sOutput + "Category: " & sTemp & "  Subcategory: " & .subcategory & sLineBreak
		'            Else
		'                sOutput = sOutput + "Category: " & sTemp & sLineBreak
		'            End If
		'        End If
		'
		'        sTemp = .CopyrightDate
		'        If sTemp <> "" Then sOutput = sOutput + "Copyright (c) " + sTemp + sLineBreak
		'        sTemp = .author
		'        If sTemp <> "" Then
		'            sOutput = sOutput + "by " + sTemp
		'
		'            sTemp = .email
		'            If sTemp <> "" Then
		'                sOutput = sOutput + " " + "<" + sTemp + ">" + sLineBreak
		'            Else
		'                sOutput = sOutput + sLineBreak
		'            End If
		'       End If
		'        sTemp = .url
		'        If sTemp <> "" Then sOutput = sOutput + "http://" + sTemp + sLineBreak
		'
		'        sTemp = .VehicleDescription
		'        If sTemp <> "" Then sOutput = sOutput + sLineBreak + sTemp + sBreak
		'
		'    End With
		'
		'    'JAW 2000.06.25
		'    'change header to include doc head tags for HTML
		'    Select Case sType
		'        Case "Class HTML", "New HTML"
		'            GetHeaderOutput = "<html><head><title>" & m_oCurrentVeh.Description.Header _
		''                & ", " & m_oCurrentVeh.Description.ClassName & "-class " & _
		''                m_oCurrentVeh.Description.category & " " & m_oCurrentVeh.Description.subcategory _
		''                & "</title></head><body>" & vbCrLf & sOutput
		'        Case Else
		'            GetHeaderOutput = sOutput
		'    End Select
	End Function
	
	Private Function CHOPCHOP(ByVal s As String) As String
		Dim GooGooGaga As Integer
		Dim ScoobyDoo As Boolean
		Dim tempbyte() As Byte
		Dim bFlag As Boolean
		Dim i As Integer
		On Error GoTo errorhandler
		'//this routine mangles the Print Output if the program is not registered
		Randomize()
		tempbyte = VB6.CopyArray(ChopCheck)
		'UPGRADE_WARNING: IsEmpty was upgraded to IsNothing and has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		If (IsNothing(tempbyte) = False) And (UBound(tempbyte) - LBound(tempbyte) = UBound(gsRegNum) - LBound(gsRegNum)) Then
			For i = 1 To UBound(gsRegNum)
				If tempbyte(i) = gsRegNum(i) Then
					bFlag = True
				Else
					bFlag = False
					Exit For
				End If
			Next 
			
			If bFlag Then
				CHOPCHOP = s
				Exit Function
			ElseIf ScoobyDoo Then 
				GoTo ScoobySnack
			Else
				
			End If
		End If
		'//mangle time
		If Len(s) <= 1 Then
			CHOPCHOP = ""
		Else
			For i = 1 To Len(s)
				Mid(s, i, 1) = Chr(Int((255 - 0 + 1) * Rnd()))
			Next 
		End If
		
		'//add some fake code that never gets executed
		If GooGooGaga = -74439050 Then
			GooGooGaga = 85858859
		End If
		CHOPCHOP = s
		If ScoobyDoo Then GoTo ScoobySnack
		Exit Function
ScoobySnack: 
		'note that this never gets called because ScoobyDoo always evaluates to False
		Resume Next
errorhandler: 
	End Function
	
	Public Function ChopCheck() As Byte()
		Dim m_oCurrentVeh As Object
		Dim tempbyte() As Byte
		Dim i As Integer
		Dim j As Single
		Dim sTName As String
		Dim lngtotal As Single
		Dim sRegNumber As String
		
		ReDim tempbyte(1)
		
#If DEBUG_MODE Then
		'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression DEBUG_MODE did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
		MsgBox "modTextOutput:ChopCheck() -- Function not available in debug mode."
		Exit Function
#End If
		On Error GoTo errorhandler
		'//one of the local reg number checkers.  There will be several of these so
		' so that a hacker will have to do some serious code hacking to disable all
		' of them
		
		'here's the reg key formula
		'1- the user's reg name and key are accepted into a byte array with each
		'   letter being actually the ascii code for that letter.  Total them up
		For i = 1 To UBound(gsRegName)
			lngtotal = lngtotal + gsRegName(i)
			'at the same time total the ascii value for every even valued ascii code
			If gsRegName(i) Mod 2 = 0 Then
				lngtotal = lngtotal + gsRegName(i)
			End If
		Next 
		'2 - the RegID is actually just a modifier to prevent two people having the same
		'    name winding up with the same ID.  This ID is unique and alone can be used
		'   to identify a user.  Multiply this to the total
		lngtotal = lngtotal * gsRegID
		'3- take the ascii value of the typename of the Body and multiply that to it
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Body. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		sTName = TypeName(m_oCurrentVeh.Body) '(BODY_KEY))
		For i = 1 To Len(sTName)
			lngtotal = lngtotal * Asc(Mid(sTName, i, 1))
		Next 
		'6- take a random seed to generate the seeded random number and multiply that
		Rnd(-1)
		Randomize(9921988)
		lngtotal = lngtotal * Rnd()
		'8- return this as a byte array that we can compare with our current one
		'how do we split this up into seperate bytes? well we know our ascii values
		'must be between 48-57, 65-90 and 97-122
		'well, we can generate a random reg code based on each number in the string
		'representation using the random seed of each number
		For i = 1 To Len(Str(lngtotal))
			j = Rnd()
			If j <= 0.33 Then
				ReDim Preserve tempbyte(i)
				Rnd(-1)
				Randomize(Asc(Mid(Str(lngtotal), i, 1)))
				tempbyte(i) = Int((57 - 48 + 1) * Rnd() + 48)
				sRegNumber = sRegNumber & Chr(tempbyte(i))
			ElseIf j <= 0.66 Then 
				ReDim Preserve tempbyte(i)
				Rnd(-1)
				Randomize(Asc(Mid(Str(lngtotal), i, 1)))
				tempbyte(i) = Int((90 - 65 + 1) * Rnd() + 65)
				sRegNumber = sRegNumber & Chr(tempbyte(i))
			Else
				ReDim Preserve tempbyte(i)
				Rnd(-1)
				Randomize(Asc(Mid(Str(lngtotal), i, 1)))
				tempbyte(i) = Int((122 - 97 + 1) * Rnd() + 97)
				sRegNumber = sRegNumber & Chr(tempbyte(i))
			End If
		Next 
		
		ChopCheck = VB6.CopyArray(tempbyte)
		Exit Function
errorhandler: 
		ReDim tempbyte(1)
		ChopCheck = VB6.CopyArray(tempbyte)
	End Function
	
	Private Function GetSubassemblyOutput() As String
		Dim sOutput As String
		Dim element As Object
		On Error GoTo err_Renamed
		
		'todo: fix
		'    For Each element In m_oCurrentVeh.Components
		'        Select Case element.Datatype
		'
		'            Case Wheel, Skid, Track, Hydrofoil, Hovercraft, _
		''                Leg, Arm, AutogyroRotor, TTRotor, CARotor, MMRotor, _
		''                Wing, Mast, Superstructure, Turret, Popturret, _
		''                OpenMount, Gasbag, Pod, SolarPanel, equipmentPod
		'
		'                sOutput = sOutput + element.PrintOutput + " "
		'
		'        End Select
		'    Next
		
		GetSubassemblyOutput = sOutput
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetSubassemblyOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	
	Private Function GetCustomComponentsOutput() As String
		Dim sOutput As String
		Dim element As Object
		
		On Error GoTo err_Renamed
		'    For Each element In m_oCurrentVeh.Components
		'        If TypeOf element Is clsSimpleCustom Then
		'           sOutput = sOutput + element.PrintOutput + " "
		'        End If
		'    Next
		'
		'    GetCustomComponentsOutput = sOutput
		'Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetCustomComponentsOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	
	Private Function GetPropulsionOutput() As String
		Dim sOutput As String
		Dim element As Object
		
		'    On Error GoTo err
		'    For Each element In m_oCurrentVeh.Components
		'        Select Case element.Datatype
		'
		'            Case WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, _
		''                FlexibodyDrivetrain, TrackedDrivetrain, LegDrivetrain, _
		''                CARRotorDrivetrain, MMRRotorDrivetrain, TTRRotorDrivetrain, _
		''                OrnithopterDrivetrain, AerialPropeller, DuctedFan, PaddleWheel, _
		''                ScrewPropeller, lightScrewPropeller, DuctedPropeller, Hydrojet, _
		''                MHDTunnel, RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, _
		''                WhiffletreeHarness, MagLevLifter, Turbojet, Turbofan, Ramjet, _
		''                TurboRamjet, Hyperfan, FusionAirRam, StandardThruster, _
		''                SuperThruster, MegaThruster, LiquidFuelRocket, MOXRocket, _
		''                IonDrive, FissionRocket, FusionRocket, OptimizedFusion, _
		''                AntimatterThermal, AntimatterPion, RowingPositions, ForeandAftRig, _
		''                SquareRig, FullRig, AerialSail, AerialSailForeAftRig, lightSail, SolidRocketEngine, _
		''                OrionEngine, TeleportationDrive, Hyperdrive, JumpDrive, _
		''                WarpDrive, QuantumConveyor, SubQuantumConveyor, _
		''                TwoQuantumConveyor
		'
		'                sOutput = sOutput + element.PrintOutput + " "
		'
		'        End Select
		'    Next
		'
		'    GetPropulsionOutput = sOutput
		'Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetPropulsionOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	
	Private Function GetAerostaticLiftOutput() As String
		Dim sOutput As String
		Dim element As Object
		
		On Error GoTo err_Renamed
		'    For Each element In m_oCurrentVeh.Components
		'        Select Case element.Datatype
		'
		'            Case ContraGravGenerator, HotAir, Hydrogen, Helium
		'
		'                sOutput = sOutput + element.PrintOutput + " "
		'
		'        End Select
		'    Next
		
		GetAerostaticLiftOutput = sOutput
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetAerostaticLiftOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	
	Private Function GetWeaponryOutput() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim element As Object
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				
				Case StoneThrower, BoltThrower, RepeatingBoltThrower, MuzzleLoader, BreechLoader, ManualRepeater, Revolver, MechanicalGatling, SlowAutoloader, FastAutoloader, lightAutomatic, HeavyAutomatic, ElectricGatling, BlueGreenLaser, RainbowLaser, Laser, UVLaser, IRLaser, Disruptor, ChargedParticleBeam, NeutralParticleBeam, Flamer, Screamer, Stunner, ParalysisBeam, XRayLaser, FusionBeam, GravityBeam, AntiparticleBeam, Graser, Disintegrator, Displacer, BeamedPowerTransmitter, MilitaryParalysisBeam, EnergyDrill, IronBomb, RetardedBomb, SmartBomb, SelfDestructSystem, ContactMine, ProximityMine, PressureTriggerMine, CommandTriggerMine, SmartTriggerMine, UnGuidedMissile, UnGuidedTorpedo, GuidedMissile, GuidedTorpedo, FlameThrower, WaterCannon, DisposableLauncher, MuzzleloadingLauncher, BreechloadingLauncher, ManualRepeaterLauncher, SlowAutoLoaderLauncher, FastAutoLoaderLauncher, RevolverLauncher, lightAutomaticLauncher, HeavyAutomaticLauncher
					
					'UPGRADE_WARNING: Couldn't resolve default property of object element.PrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sOutput = CStr(CDbl(sOutput) + element.PrintOutput + CDbl(" "))
					
			End Select
		Next element
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			If TypeOf element Is clsWeaponAmmunition Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.PrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sOutput = CStr(CDbl(sOutput) + element.PrintOutput + CDbl(" "))
				
			End If
		Next element
		
		GetWeaponryOutput = sOutput
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetWeaponryOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	
	Private Function GetWeaponAccessoriesOutput() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim element As Object
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				
				Case PartialStabilizationGear, FullStabilizationGear, UniversalMount, CasemateMount, DoorMount, Cyberslave, AntiBlastMagazine
					
					'UPGRADE_WARNING: Couldn't resolve default property of object element.PrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sOutput = CStr(CDbl(sOutput) + element.PrintOutput + CDbl(" "))
					
			End Select
		Next element
		
		GetWeaponAccessoriesOutput = sOutput
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetWeaponAccessoriesOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	
	Private Function GetWeaponLinksOutput() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim element As Object
		Dim sKeyArray() As String
		Dim i As Integer
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			If TypeOf element Is clsWeaponLink Then
				'//append the weapons that are in the link
				'UPGRADE_WARNING: Couldn't resolve default property of object element.getcurrentkeys. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sKeyArray = VB6.CopyArray(element.getcurrentkeys)
				If sKeyArray(1) = "" Then
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Key. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sOutput = CDbl(sOutput) + element.Key & " controls "
					For i = 1 To UBound(sKeyArray)
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sOutput = CDbl(sOutput) + m_oCurrentVeh.Components(sKeyArray(i)).Description & ", "
					Next 
					'//delete the last "," and replace it with "."
					sOutput = Left(sOutput, Len(sOutput) - 2)
					sOutput = sOutput & ".  "
					
				End If
				
			End If
		Next element
		
		GetWeaponLinksOutput = sOutput
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetWeaponLinksOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	Private Function GetCommunicationsOutput() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim element As Object
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				
				Case RadioDirectionFinder, RadioCommunicator, TightBeamRadio, VLFRadio, CellularPhone, CellularPhonewithRadio, RadioJammer, ElfReceiver, LaserCommunicator, NeutrinoCommunicator, GravityRippleCommunicator
					
					'UPGRADE_WARNING: Couldn't resolve default property of object element.PrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sOutput = CStr(CDbl(sOutput) + element.PrintOutput + CDbl(" "))
					
			End Select
		Next element
		
		GetCommunicationsOutput = sOutput
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetCommunicationsOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	Private Function GetSensorsOutput() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim element As Object
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				
				Case Headlight, Searchlight, InfraredSearchlight, AstronomicalInstruments, Telescope, lightAmplification, LowlightTV, ExtendableSensorPeriscope, Radar, Ladar, NavigationalRadar, AntiCollisionRadar, AESA, LowResImagingRadar, HiResImagingRadar, ActiveSonar, PassiveSonar, PassiveInfrared, Thermograph, PassiveRadar, PESA, Geophone, MAD, MultiScanner, ChemScanner, RadScanner, BioScanner, GravScanner, RangingSoundDetector, SurveillanceSoundDetector, MeteorologicalInstruments, LowResPlanetarySurveyArray, MedResPlanetarySurveyArray, HighResPlanetarySurveyArray
					
					'UPGRADE_WARNING: Couldn't resolve default property of object element.PrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sOutput = CStr(CDbl(sOutput) + element.PrintOutput + CDbl(" "))
					
			End Select
		Next element
		
		GetSensorsOutput = sOutput
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetSensorsOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	Private Function GetAudioVisualOutput() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim element As Object
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				
				Case SoundSystem, FlightRecorder, VehicleCamera, DigitalVehicleCamera, ReconCamera, DigitalReconCamera
					
					'UPGRADE_WARNING: Couldn't resolve default property of object element.PrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sOutput = CStr(CDbl(sOutput) + element.PrintOutput + CDbl(" "))
					
			End Select
		Next element
		
		GetAudioVisualOutput = sOutput
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetAudioVisualOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	
	Private Function GetNavigationOutput() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim element As Object
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				
				Case NavigationInstruments, AutoPilot, IFF, Transponder, INS, GPS, MilitaryGPS, TFR
					
					'UPGRADE_WARNING: Couldn't resolve default property of object element.PrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sOutput = CStr(CDbl(sOutput) + element.PrintOutput + CDbl(" "))
					
			End Select
		Next element
		
		GetNavigationOutput = sOutput
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetNavigationOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	
	Private Function GetTargetingOutput() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim element As Object
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				
				Case ImprovedOpticalBombSight, AdvancedOpticalBombSight, OpticalBombSight, FireDirectionCenter, HUDWAC, PupilHUDWAC, LaserRangeFinder, LaserDesignator, LaserSpotTracker
					
					'UPGRADE_WARNING: Couldn't resolve default property of object element.PrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sOutput = CStr(CDbl(sOutput) + element.PrintOutput + CDbl(" "))
					
			End Select
		Next element
		
		GetTargetingOutput = sOutput
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetTargetingOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	
	Private Function GetECMOutput() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim element As Object
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				
				Case RadarDetector, LaserSensor, LaserRadarDetector, AreaRadarJammer, DeceptiveRadarJammer, InfraredJammer, ChaffDecoyDischarger, SmokeDecoyDischarger, FlareDecoyDischarger, SonarDecoyDischarger, HotSmokeDecoyDischarger, PrismDecoyDischarger, BlackOutGasDecoyDischarger, RadarReflector, BlipEnhancer, TEMPEST
					
					'UPGRADE_WARNING: Couldn't resolve default property of object element.PrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sOutput = CStr(CDbl(sOutput) + element.PrintOutput + CDbl(" "))
					
			End Select
		Next element
		
		GetECMOutput = sOutput
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetECMOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	
	Private Function GetComputersOutput() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim element As Object
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				
				Case MacroFrame, MainFrame, MicroFrame, MiniComputer, SmallComputer, Terminal
					
					'UPGRADE_WARNING: Couldn't resolve default property of object element.PrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sOutput = CStr(CDbl(sOutput) + element.PrintOutput + CDbl(" "))
					
			End Select
		Next element
		
		GetComputersOutput = sOutput
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetComputersOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	
	Private Function GetSoftwareOutput() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim element As Object
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			If TypeOf element Is clsSoftware Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.PrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sOutput = CStr(CDbl(sOutput) + element.PrintOutput + CDbl(" "))
			End If
		Next element
		
		
		GetSoftwareOutput = sOutput
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetSoftwareOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	
	
	Private Function GetNeuralInterfaceSystemOutput() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim element As Object
		
		On Error GoTo err_Renamed
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			If TypeOf element Is clsNeuralInterfaceSystem Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.PrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sOutput = CStr(CDbl(sOutput) + element.PrintOutput + CDbl(" "))
			End If
		Next element
		
		GetNeuralInterfaceSystemOutput = sOutput
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetNeuralInterfaceSystemOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	
	Private Function GetMiscellaneousOutput() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim element As Object
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				
				Case ArmMotor, FireExtinguisherSystem, FullFireSuppressionSystem, CompactFireSuppressionSystem, BilgePump, CompleteWorkshop, MechanicWorkshop, EngineeringWorkshop, ElectronicsWorkshop, ArmouryWorkshop, CompleteMiniWorkshop, ScienceLab, MiniMechanicWorkshop, MiniElectronicsWorkshop, MiniEngineeringWorkshop, MiniArmouryWorkshop
					
					'UPGRADE_WARNING: Couldn't resolve default property of object element.PrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sOutput = CStr(CDbl(sOutput) + element.PrintOutput + CDbl(" "))
					
				Case ExtendableLadder, Crane, Winch, PowerShovel, WreckingCrane, ForkLift, VehicularBridge, LaunchCatapult, SkyHook, Bore, SuperBore, EnergyDrill, TractorBeam, PressorBeam, CombinationBeam, CraneWithElectroMagnet, OperatingRoom, StretcherPallet, EmergencySupportUnit, EmergencylightsandSiren, CryonicCapsule, Automed, DiagnosisTable, Stage, Hall, BarRoom, ConferenceRoom, MovieScreenandProjector, MovieScreenandProjectorSmall, HoloventureZone, CargoRamp, Airlock, MembraneAirlock, Forcelock, PassageTube, ArmoredPassageTube, TeleportProjector, BrigsandRestraints, BurglarAlarm, HighSecurityAlarm, MutableLicensePlate, OilSprayer, PaintSprayer, SmokeScreen, SpikeDropper, VehicleBay, HangerBay, DryDock, SpaceDock, ExternalCradle, ArrestorHook, VehicularParachute, RefuellingProbe, RefuellingDrogue, FuelElectrolysisSystem, HydrogenFuelScoop, AtmosphereProcessor, NuclearDamper, SmallRealityStabilizer, MediumRealityStabilizer, HeavyRealityStabilizer, ModularSocket, Module_Renamed
					
					'UPGRADE_WARNING: Couldn't resolve default property of object element.PrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sOutput = CStr(CDbl(sOutput) + element.PrintOutput + CDbl(" "))
					
			End Select
		Next element
		
		GetMiscellaneousOutput = sOutput
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetMiscellaneousOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	
	Private Function GetVehicleControlsOutput() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim element As Object
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				
				Case PrimitiveManeuverControl, ElectronicDivingControl, ComputerizedDivingControl, MechanicalManeuverControl, ElectronicManeuverControl, ComputerizedManeuverControl, MechanicalDivingControl
					
					'UPGRADE_WARNING: Couldn't resolve default property of object element.PrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sOutput = CStr(CDbl(sOutput) + element.PrintOutput + CDbl(" "))
					
			End Select
		Next element
		
		GetVehicleControlsOutput = sOutput
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetVehicleControlsOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	
	Private Function GetCrewStationsOutput() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim element As Object
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				
				Case CrampedCrewStation, NormalCrewStation, RoomyCrewStation, CycleCrewStation, HarnessCrewStation
					
					
					'UPGRADE_WARNING: Couldn't resolve default property of object element.PrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sOutput = CStr(CDbl(sOutput) + element.PrintOutput + CDbl(" "))
					
			End Select
		Next element
		
		GetCrewStationsOutput = sOutput
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetCrewStationsOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	
	Private Function GetOccupancyOutput() As String
		Dim m_oCurrentVeh As Object
		
		Dim sOutput As String
		Dim element As Object
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With m_oCurrentVeh.crew
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .numshifts > 1 Then sOutput = NumericToString(.numshifts) & " shifts. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .numcaptains > 0 Then sOutput = sOutput & NumericToString(.numcaptains) & " captains. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .NumOfficers > 0 Then sOutput = sOutput & NumericToString(.NumOfficers) & " officers. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .NumCrewStationOperators > 0 Then sOutput = sOutput & NumericToString(.NumCrewStationOperators) & " crew station operators. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .NumWeaponLoaders > 0 Then sOutput = sOutput & NumericToString(.NumWeaponLoaders) & " weapon loaders. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .NumRowers > 0 Then sOutput = sOutput & NumericToString(.NumRowers) & " rowers. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .NumSailors > 0 Then sOutput = sOutput & NumericToString(.NumSailors) & " sailors. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .NumRiggers > 0 Then sOutput = sOutput & NumericToString(.NumRiggers) & " sail riggers. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .NumFuelStokers > 0 Then sOutput = sOutput & NumericToString(.NumFuelStokers) & " fuel stokers. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .NumMechanics > 0 Then sOutput = sOutput & NumericToString(.NumMechanics) & " mechanics. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .NumServiceCrewmen > 0 Then sOutput = sOutput & NumericToString(.NumServiceCrewmen) & " service crewmen. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .NumMedics > 0 Then sOutput = sOutput & NumericToString(.NumMedics) & " medics. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .NumScientists > 0 Then sOutput = sOutput & NumericToString(.NumScientists) & " scientists. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .NumAuxiliaryVehicleCrew > 0 Then sOutput = sOutput & NumericToString(.NumAuxiliaryVehicleCrew) & " auxiliary vehicle crewmen. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .NumStewards > 0 Then sOutput = sOutput & NumericToString(.NumStewards) & " stewards. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .NumLuxury > 0 Then sOutput = sOutput & NumericToString(.NumLuxury) & " luxury class passengers. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .NumFirstClass > 0 Then sOutput = sOutput & NumericToString(.NumFirstClass) & " first class passengers. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .NumSecondClass > 0 Then sOutput = sOutput & NumericToString(.NumSecondClass) & " second class passengers. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .NumSteerage > 0 Then sOutput = sOutput & NumericToString(.NumSteerage) & " steerage passengers. "
			
			' append whether its short or long Occupancy
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sOutput = .Occupancy & ". " & sOutput
		End With
		
		GetOccupancyOutput = sOutput
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetOccupancyOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	
	Private Function GetAccomodationsOutput() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim element As Object
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				
				Case CrampedSeat, NormalSeat, RoomySeat, CrampedStandingRoom, NormalStandingRoom, RoomyStandingRoom, CycleSeat, Hammock, Bunk, Cabin, LuxuryCabin, Suite, LuxurySuite, SmallGalley
					
					
					'UPGRADE_WARNING: Couldn't resolve default property of object element.PrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sOutput = CStr(CDbl(sOutput) + element.PrintOutput + CDbl(" "))
					
			End Select
		Next element
		
		GetAccomodationsOutput = sOutput
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetAccomodationsOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	
	Private Function GetEnvironmentalSystemsOutput() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim element As Object
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				
				Case TotalLifeSystem, ArtificialGravityUnit, EnvironmentalControl, NBCKit, LimitedLifeSystem, FullLifeSystem
					
					'UPGRADE_WARNING: Couldn't resolve default property of object element.PrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sOutput = CStr(CDbl(sOutput) + element.PrintOutput + CDbl(" "))
			End Select
		Next element
		
		'append the provisions to it
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If element.Datatype = Provisions Then
				'UPGRADE_WARNING: Couldn't resolve default property of object element.PrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sOutput = CStr(CDbl(sOutput) + element.PrintOutput + CDbl(" "))
			End If
		Next element
		GetEnvironmentalSystemsOutput = sOutput
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:EnvironmentalSystemsOutput -- Error #" & Err.Number & " " & Err.Description)
	End Function
	
	Private Function GetSafetySystemsOutput() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim element As Object
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				Case EjectionSeat, CrewEscapeCapsule, Airbag, CrashWeb, WombTank, GravityWeb, GravCompensator
					
					'UPGRADE_WARNING: Couldn't resolve default property of object element.PrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sOutput = CStr(CDbl(sOutput) + element.PrintOutput + CDbl(" "))
			End Select
		Next element
		GetSafetySystemsOutput = sOutput
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetSafetySystemsOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	
	Public Function GetPowerSystemsOutPut() As String
		Dim m_oCurrentVeh As Object
		'UPGRADE_ISSUE: clsProfilePower object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim oProfile As clsProfilePower
		Dim oGroup As clsSupplyConsumeGroup
		Dim iGroupCount As Integer
		Dim i As Integer
		Dim iSupplierCount As Integer
		Dim iConsumerCount As Integer
		Dim sTemp As String
		Dim j As Integer
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Profiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each oProfile In m_oCurrentVeh.Profiles
			'UPGRADE_WARNING: Couldn't resolve default property of object oProfile.Key. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sTemp = sTemp & "Profile " & oProfile.Key & vbNewLine
			'UPGRADE_WARNING: Couldn't resolve default property of object oProfile.groupcount. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			iGroupCount = oProfile.groupcount
			If iGroupCount > 0 Then
				For i = 1 To iGroupCount
					'UPGRADE_WARNING: Couldn't resolve default property of object oProfile.Group. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					oGroup = oProfile.Group(i)
					' get the suppliers
					sTemp = sTemp & " Suppliers " & vbNewLine
					iSupplierCount = oGroup.SupplierCount
					For j = 1 To iSupplierCount
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sTemp = sTemp & m_oCurrentVeh.Components(oGroup.Supplier(j)).Description
					Next 
					' get the consumers
					sTemp = sTemp & " Consumers " & vbNewLine
					iConsumerCount = oGroup.ConsumerCount
					For j = 1 To iConsumerCount
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sTemp = sTemp & m_oCurrentVeh.Components(oGroup.consumer(j)).Description
					Next 
					sTemp = sTemp
				Next 
			End If
			
			GetPowerSystemsOutPut = sTemp
			' get the keys for each supplier in each group
			
			
			' get the keys for all consumers attached to each group
			
			
		Next oProfile ' next profile
		'UPGRADE_NOTE: Object oGroup may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oGroup = Nothing
		'UPGRADE_NOTE: Object oProfile may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oProfile = Nothing
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetPowerSystemsOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	'Private Function GetPowerSystemsOutPut() As String
	'  MPJ 05/27/02 ENTIRE FUNCTION OBSOLETE and NON FUNCTIONAL WITH NEW POWER SYSTEM PROFILES
	'    Dim sOutput As String
	'    Dim element As Object
	'    Dim sKeyArray() As String
	'    Dim i As Long
	'    Dim sngPowerRemaining As Single
	'
	'    'todo: this is simplified because we can get this info directly from the
	'    ' m_SC_Groups in each profile
	'
	'    ' todo: however, we do need seperate writes ups for each profile
	'
	'    For Each element In m_oCurrentVeh.Components
	'        Select Case element.Datatype
	'
	'            Case MuscleEngine, GasolineEngine, HPGasolineEngine, _
	''                TurboGasolineEngine, SuperGasolineEngine, TurboHPGasolineEngine, _
	''                SuperHPGasolineEngine, StandardDieselEngine, _
	''                TurboStandardDieselEngine, MarineDieselEngine, _
	''                HPDieselEngine, TurboHPDieselEngine, CeramicEngine, _
	''                TurboCeramicEngine, SuperCeramicEngine, HPCeramicEngine, _
	''                TurboHPCeramicEngine, SuperHPCeramicEngine, _
	''                HydrogenCombustionEngine, EarlySteamEngine, _
	''                ForcedDraftSteamEngine, TripleExpansionSteamEngine, SteamTurbine
	'
	'                sOutput = sOutput + element.PrintOutput + " "
	'
	'                '//append the systems that it powers
	'                sKeyArray = element.GetCurrentConsumptionSystemKeys
	'                If sKeyArray(1) = "" Then
	'                Else
	'                    sOutput = sOutput + "Powers the "
	'                    For i = 1 To UBound(sKeyArray)
	'                        sOutput = sOutput + m_oCurrentVeh.Components(sKeyArray(i)).Description & ", "
	'                    Next
	'                    '//delete the last "," and replace it with "."
	'                    sOutput = Left(sOutput, Len(sOutput) - 2)
	'                    sngPowerRemaining = element.Output - element.PowerConsumed
	'                    If sngPowerRemaining > 0 Then
	'                        sOutput = sOutput + " with " & sngPowerRemaining & " kW in reserve."
	'                    Else
	'                        sOutput = sOutput + ".  "
	'                    End If
	'                End If
	'
	'            Case StandardGasTurbine, HPGasTurbine, OptimizedGasTurbine, _
	''                StandardMHDTurbine, HPMHDTurbine, FuelCell, FissionReactor, _
	''                RTGReactor, NPU, FusionReactor, AntimatterReactor, _
	''                TotalConversionPowerPlant, CosmicPowerPlant, Soulburner, _
	''                ElementalFurnace, ManaEngine, Carnivore, Herbivore, Omnivore, _
	''                Vampire, ClockWork, LeadAcidBattery, AdvancedBattery, _
	''                Flywheel, RechargeablePowerCell, PowerCell, Snorkel, _
	''                ElectricContactPower, LaserBeamedPowerReceiver, _
	''                MaserBeamedPowerReceiver, SolarCellArray
	'
	'                sOutput = sOutput + element.PrintOutput + " "
	'
	'                '//append the systems that it powers
	'                sKeyArray = element.GetCurrentConsumptionSystemKeys
	'                If sKeyArray(1) = "" Then
	'                Else
	'                    sOutput = sOutput + "Powers the "
	'                    For i = 1 To UBound(sKeyArray)
	'                        sOutput = sOutput + m_oCurrentVeh.Components(sKeyArray(i)).Description & ", "
	'                    Next
	'                    '//delete the last "," and replace it with "."
	'                    sOutput = Left(sOutput, Len(sOutput) - 2)
	'                    sngPowerRemaining = element.Output - element.PowerConsumed
	'                    If sngPowerRemaining > 0 Then
	'                        sOutput = sOutput + " with " & sngPowerRemaining & " kW in reserve."
	'                    Else
	'                        sOutput = sOutput + ".  "
	'                    End If
	'                End If
	'        End Select
	'
	'    Next
	'
	'
	'
	'    GetPowerSystemsOutPut = sOutput
	'End Function
	
	Private Function GetFuelOutput() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim element As Object
		
		'TODO: Is this where endurance gets spit out?
		'
		' 'find the endurance of the engine
		'    mvarEndurance = 0 'reset the variable
		'    If mvarFuelStorageKeyChain(1) <> "" Then
		'        For i = 1 To UBound(mvarFuelStorageKeyChain)
		'            mvarEndurance = mvarEndurance + Veh.Components(mvarFuelStorageKeyChain(i)).capacity
		'        Next
		'        mvarEndurance = mvarEndurance / mvarFuelConsumption
		'    Else
		'        mvarEndurance = 0
		'    End If
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				Case AntiMatterBay, CoalBunker, WoodBunker, StandardTank, lightTank, UltralightTank, StandardSelfSealingTank, lightSelfSealingTank, UltralightSelfSealingTank
					
					'UPGRADE_WARNING: Couldn't resolve default property of object element.PrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sOutput = CStr(CDbl(sOutput) + element.PrintOutput + CDbl(" "))
			End Select
		Next element
		GetFuelOutput = sOutput
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetFuelOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	
	Private Function GetSurfaceAreaOutput() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim element As Object
		Dim totalsurfacearea As Single
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				Case Body, Wheel, Skid, Track, Hydrofoil, Hovercraft, Leg, Arm, AutogyroRotor, TTRotor, CARotor, MMRotor, Wing, Mast, Superstructure, Turret, Popturret, OpenMount, Gasbag, Pod
					
					'UPGRADE_WARNING: Couldn't resolve default property of object element.SurfaceArea. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					totalsurfacearea = totalsurfacearea + element.SurfaceArea
					'UPGRADE_WARNING: Couldn't resolve default property of object element.SurfaceArea. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.abbrev. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sOutput = CStr(CDbl(sOutput) + element.abbrev + CDbl(" ") + CDbl(VB6.Format(element.SurfaceArea, Settings.FormatString)) + CDbl(". "))
			End Select
		Next element
		GetSurfaceAreaOutput = sOutput & "total " & VB6.Format(totalsurfacearea, Settings.FormatString) & "."
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetSurfaceAreaOutput -- Error #" & Err.Number & " " & Err.Description)
	End Function
	
	Private Function GetStructureOutput() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim element As Object
		Dim sBodyStruct As String
		Dim tOutput As String
		
		On Error GoTo err_Renamed
		'get the structure of the body first
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With m_oCurrentVeh.Components(BODY_KEY)
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sBodyStruct = CStr(element.Description + " - " + .FrameStrength + CDbl(" frame") + CDbl(" with ") + .Materials + CDbl(" materials. "))
		End With
		
		sOutput = sBodyStruct
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				'note Open Mount, Mast and Gasbag are not included here
				Case Body, Wheel, Skid, Track, Hydrofoil, Hovercraft, Leg, Arm, AutogyroRotor, TTRotor, CARotor, MMRotor, Wing, Superstructure, Turret, Popturret, Pod
					
					With element
						'UPGRADE_WARNING: Couldn't resolve default property of object element.Materials. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object element.FrameStrength. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object element.abbrev. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						tOutput = CStr(element.abbrev + " - " + .FrameStrength + CDbl(" frame") + CDbl(" with ") + .Materials + CDbl(" materials. "))
					End With
					
					'only print this if its different than the Body's structure
					If tOutput <> sBodyStruct Then
						sOutput = sOutput & " " & tOutput
					End If
			End Select
		Next element
		
		GetStructureOutput = sOutput
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetStructureOutput -- Error #" & Err.Number & " " & Err.Description)
		Resume Next
	End Function
	Private Function GetHitPointsOutput() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim element As Object
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				Case Body, Wheel, Skid, Track, Hydrofoil, Hovercraft, Leg, Arm, AutogyroRotor, TTRotor, CARotor, MMRotor, Wing, Mast, Superstructure, Turret, Popturret, OpenMount, Gasbag, Pod, equipmentPod, SolarPanel
					
					'UPGRADE_WARNING: Couldn't resolve default property of object element.HitPoints. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.abbrev. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sOutput = CStr(CDbl(sOutput) + element.abbrev + CDbl(" ") + CDbl(VB6.Format(element.HitPoints)) + CDbl(", "))
			End Select
		Next element
		GetHitPointsOutput = Left(sOutput, Len(sOutput) - 2) & "."
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetHitPointsOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	
	Private Function GetArmorOutput() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim element As Object
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				Case ArmorComplexFacing, ArmorBasicFacing, ArmorOpenFrame, ArmorGunShield, ArmorLocation, ArmorComponent, ArmorOverall, ArmorWheelGuard
					
					'UPGRADE_WARNING: Couldn't resolve default property of object element.PrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.LogicalParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sOutput = CStr(CDbl(sOutput) + m_oCurrentVeh.Components(element.LogicalParent).CustomDescription + CDbl(" armor: ") + element.PrintOutput + CDbl(vbNewLine))
			End Select
		Next element
		GetArmorOutput = sOutput
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetArmorOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	
	Private Function GetStatisticsOutput() As String
		Dim m_oCurrentVeh As Object
		Dim sTemp As String
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With m_oCurrentVeh.Stats
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sTemp = "Empty weight " & VB6.Format(.EmptyWeight, Settings.FormatString) & " lbs., "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .UsualInternalPayload <> 0 Then sTemp = sTemp & "Internal payload " & VB6.Format(.UsualInternalPayload, Settings.FormatString) & " lbs., "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sTemp = sTemp & "Loaded weight " & VB6.Format(.LoadedWeight, Settings.FormatString) & " lbs., "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .SubmergedWeight <> 0 Then sTemp = sTemp & "Submerged weight " & VB6.Format(.SubmergedWeight, Settings.FormatString) & " lbs., "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sTemp = sTemp & "Volume " & VB6.Format(.TotalVolume, Settings.FormatString) & " cf. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sTemp = sTemp & "Size modifier " & VB6.Format(.SizeModifier) & ". "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sTemp = sTemp & "Cost $" & VB6.Format(.TotalPrice, Settings.FormatString) & ". "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Stats. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sTemp = sTemp & "HT " & VB6.Format(.StructuralHealth)
		End With
		GetStatisticsOutput = sTemp
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetStatisticsOutput -- Error #" & Err.Number & " " & Err.Description)
		
	End Function
	
	Private Function GetSpaceOutput() As String
		Dim m_oCurrentVeh As Object
		'access, empty and cargo space
		Dim sAccessOutput As String
		Dim sEmptyOutput As String
		Dim sCargoOutput As String
		Dim element As Object
		Dim sOutput As String
		
		On Error GoTo err_Renamed
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				
				Case Body, Wheel, Skid, Track, Hydrofoil, Hovercraft, Leg, Arm, AutogyroRotor, TTRotor, CARotor, MMRotor, Wing, Mast, Superstructure, Turret, Popturret, OpenMount, Gasbag, Pod, SolarPanel, equipmentPod
					
					'UPGRADE_WARNING: Couldn't resolve default property of object element.EmptySpace. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If element.EmptySpace <> 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object element.EmptySpace. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object element.abbrev. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sEmptyOutput = CStr(CDbl(sEmptyOutput) + element.abbrev + CDbl(" ") + CDbl(VB6.Format(element.EmptySpace, Settings.FormatString)) + CDbl(" cf, "))
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object element.AccessSpace. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If element.AccessSpace <> 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object element.AccessSpace. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object element.abbrev. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sAccessOutput = CStr(CDbl(sAccessOutput) + element.abbrev + CDbl(" ") + CDbl(VB6.Format(element.AccessSpace, Settings.FormatString)) + CDbl(" cf, "))
					End If
				Case Cargo
					'UPGRADE_WARNING: Couldn't resolve default property of object element.PrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sCargoOutput = CStr(CDbl(sCargoOutput) + element.PrintOutput + CDbl(" "))
					
			End Select
		Next element
		sCargoOutput = Left(sCargoOutput, Len(sCargoOutput) - 1)
		If sCargoOutput <> "" Then
			sOutput = "Space: " & sCargoOutput
		End If
		If sAccessOutput <> "" Then
			sAccessOutput = Left(sAccessOutput, Len(sAccessOutput) - 2)
			sAccessOutput = "(" & sAccessOutput & ")"
			If sOutput = "" Then sOutput = "Space: "
			sOutput = sOutput & " Access space " & sAccessOutput & "."
		End If
		If sEmptyOutput <> "" Then
			sEmptyOutput = Left(sEmptyOutput, Len(sEmptyOutput) - 2)
			sEmptyOutput = "(" & sEmptyOutput & ")"
			If sOutput = "" Then sOutput = "Space: "
			sOutput = sOutput & " Empty space " & sEmptyOutput & "."
		End If
		GetSpaceOutput = sOutput
		'this will error if the component doesnt have Access or Emtpyspace properties.
		'so will just resume past the error
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetSpaceOutput -- Error #" & Err.Number & " " & Err.Description)
		Resume Next
	End Function
	
	Private Function GetStructuralOptionsOutput() As String
		Dim m_oCurrentVeh As Object
		
		Dim element As Object
		Dim sOutput As String
		Dim bControlledInstability As Boolean
		Dim bImprovedSuspension As Boolean
		
		On Error GoTo err_Renamed
		'//get the structural options that are stored in the body
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With m_oCurrentVeh.Options
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .RollStabilizers Then
				
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If m_oCurrentVeh.surface.Submersible = True Then
				
			End If
		End With
		'//now get the rest of the structural options from the various other subsassemblies
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				Case Body, Superstructure, Popturret, Turret
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Compartmentalization. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If element.Compartmentalization <> "none" Then
						'compartmentalization
						'UPGRADE_WARNING: Couldn't resolve default property of object element.abbrev. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object element.Compartmentalization. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sOutput = sOutput & " " & element.Compartmentalization & " compartmentalization in " & element.abbrev & "."
					End If
				Case Wing
					'folding wings or rotors
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Folding. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If element.Folding Then
						sOutput = sOutput & " Folding wings."
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object element.VariableSweep. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If element.VariableSweep <> "none" Then
						'UPGRADE_WARNING: Couldn't resolve default property of object element.VariableSweep. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sOutput = sOutput & element.VariableSweep & " variable sweep wings."
					End If
					'controlled instability
					'UPGRADE_WARNING: Couldn't resolve default property of object element.ControlledInstability. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (element.ControlledInstability) And (bControlledInstability = False) Then
						sOutput = sOutput & " Controlled instability."
						bControlledInstability = True
					End If
				Case TTRotor, AutogyroRotor, MMRotor, CARotor
					'folding wings or rotors
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Folding. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If element.Folding Then
						sOutput = sOutput & " Folding rotors."
					End If
					'controlled instability
					'UPGRADE_WARNING: Couldn't resolve default property of object element.ControlledInstability. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (element.ControlledInstability) And (bControlledInstability = False) Then
						sOutput = sOutput & " Controlled instability."
						bControlledInstability = True
					End If
				Case Track, Skid, Leg
					'UPGRADE_WARNING: Couldn't resolve default property of object element.ImprovedSuspension. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (element.ImprovedSuspension) And (bImprovedSuspension = False) Then
						sOutput = sOutput & " Improved Suspension."
						bImprovedSuspension = True
					End If
					
				Case Wheel
					'UPGRADE_WARNING: Couldn't resolve default property of object element.ImprovedSuspension. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (element.ImprovedSuspension) And (bImprovedSuspension = False) Then
						sOutput = sOutput & " Improved Suspension."
						bImprovedSuspension = True
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object element.ImprovedBrakes. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If element.ImprovedBrakes Then
						sOutput = sOutput & " Improved Brakes."
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object element.AllwheelSteering. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If element.AllwheelSteering Then
						sOutput = sOutput & " All wheel steering."
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Smartwheels. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If element.Smartwheels Then
						sOutput = sOutput & " Smart wheels."
					End If
			End Select
		Next element
		
		If sOutput <> "" Then sOutput = "Structural Options: " & sOutput
		
		GetStructuralOptionsOutput = sOutput
		Exit Function
err_Renamed: 
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Debug.Print("modTextOutput:GetStructuralOptionsOutput --  Error #" & Err.Number & " " & TypeName(element) & " " & Err.Description)
		Resume Next
	End Function
	
	Private Function GetBodyFeatures() As String
		Dim sOutput As String
		Dim element As Object
		Dim sSlope As String
		Dim sOutput2 As String
		
		'On Error GoTo err
		'    With m_oCurrentVeh.surface
		'        If .FloatationHull Then
		'            sOutput = sOutput & " Floatation hull " ' todo: fix (rating " & m_oCurrentVeh.stats.FloatationRating & " lbs)."
		'        End If
		'        If .SubmarineLines Then
		'            sOutput = sOutput & " Submarine lines."
		'        End If
		'        If .HydrodynamicLines <> "none" Then
		'            sOutput = sOutput & " " & .HydrodynamicLines & " hydrodynamic lines."
		'        End If
		'        If .Catamaran Then
		'            sOutput = sOutput & " Catamaran."
		'        End If
		'        If .Trimaran Then
		'            sOutput = sOutput & " Trimaran."
		'        End If
		'        If .StreamLining <> "none" Then
		'            sOutput = sOutput & " " & .StreamLining & " streamlining."
		'        End If
		'    End With
		'    '//to get the slope i must check the Body, Superstructure and Turrets and Popturrets
		'    For Each element In Vehicle
		'        Select Case element.Datatype
		'            Case Body, Superstructure, Turret, Popturret
		'                sSlope = ""
		'                If element.slopef <> "none" Then
		'                    If sSlope <> "" Then
		'                        sSlope = sSlope & ", front " & element.slopef
		'                    Else
		'                        sSlope = sSlope & "front " & element.slopef
		'                    End If
		'                End If
		'                If element.slopeb <> "none" Then
		'                    If sSlope <> "" Then
		'                        sSlope = sSlope & ", back " & element.slopeb
		'                    Else
		'                        sSlope = sSlope & "back " & element.slopeb
		'                    End If
		'                End If
		'                If element.slopel <> "none" Then
		'                    If sSlope <> "" Then
		'                        sSlope = sSlope & ", left " & element.slopel
		'                    Else
		'                        sSlope = sSlope & "left " & element.slopel
		'                    End If
		'                End If
		'                If element.SlopeR <> "none" Then
		'                    If sSlope <> "" Then
		'                        sSlope = sSlope & ", right " & element.SlopeR
		'                    Else
		'                        sSlope = sSlope & "right " & element.SlopeR
		'                    End If
		'                End If
		'                If sSlope <> "" Then
		'                    sOutput2 = sOutput2 + " Slope on " & element.abbrev & ": " & sSlope & "."
		'                End If
		'        End Select
		'    Next
		'
		'    sOutput = sOutput + sOutput2
		'    If sOutput <> "" Then sOutput = sOutput
		'    GetBodyFeatures = sOutput
		'    Exit Function
		'err:
		'    Debug.Print "modTextOutput:GetBodyFeatures --  Error #" & err.Number & " " & TypeName(element) & " " & err.Description
	End Function
	
	Private Function GetSurfaceFeaturesOutput() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim element As Object
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With m_oCurrentVeh.surface
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .Sealed Then sOutput = sOutput & "sealed. "
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .Sealed = False Then
				'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If .WaterProof Then
					sOutput = sOutput & "waterproofed. "
				End If
			End If
			
			'concealment and stealth
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .Camouflage Then sOutput = sOutput & "camouflage. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .infraredcloaking <> "none" Then sOutput = CStr(CDbl(sOutput) + .infraredcloaking + CDbl(" infrared cloaking. "))
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .EmissionCloaking <> "none" Then sOutput = CStr(CDbl(sOutput) + .EmissionCloaking + CDbl(" emission cloaking. "))
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .SoundBaffling <> "none" Then sOutput = CStr(CDbl(sOutput) + .SoundBaffling + CDbl(" sound baffling. "))
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .stealth <> "none" Then sOutput = CStr(CDbl(sOutput) + .stealth + CDbl(" stealth. "))
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .LiquidCrystal Then sOutput = sOutput & "liquid crystal skin. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .Chameleon <> "none" Then sOutput = CStr(CDbl(sOutput) + .Chameleon + CDbl(" chameleon system. "))
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .PsiShielding Then sOutput = sOutput & "Psi Shielding. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If m_oCurrentVeh.Components(BODY_KEY).liftingbody Then sOutput = sOutput & "Lifting Body. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If m_oCurrentVeh.Components(BODY_KEY).FlexibodyOption Then sOutput = sOutput & "Flexibody. "
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For	Each element In m_oCurrentVeh.Components
				'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Select Case element.Datatype
					Case ForceScreen, DeflectorField, VariableForceScreen
						'UPGRADE_WARNING: Couldn't resolve default property of object element.PrintOutput. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sOutput = CStr(CDbl(sOutput) + element.PrintOutput)
						
				End Select
			Next element
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .bMagicLevitation Then sOutput = sOutput & "Magic Levitation. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .bAntigravityCoating Then sOutput = sOutput & "Antigravity Coating. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .bSuperScienceCoating Then sOutput = sOutput & "Super Science Coating. "
		End With
		
		If sOutput <> "" Then sOutput = "Surface Features: " & sOutput
		GetSurfaceFeaturesOutput = sOutput
		Exit Function
err_Renamed: 
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Debug.Print("modTextOutput:GetSufaceFeaturesOutput --  Error #" & Err.Number & " " & TypeName(element) & " " & Err.Description)
	End Function
	
	Private Function GetOtherSurfaceFeatures() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim element As Object
		Dim sDefensive As String
		Dim sOutput2 As String
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With m_oCurrentVeh.Options
			'other
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .Convertible <> "none" Then sOutput = CStr(CDbl(sOutput) + .Convertible + CDbl(". "))
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .Ram Then sOutput = sOutput & "Ram. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .Bulldozer Then sOutput = sOutput & "Bulldozer. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .Plow Then sOutput = sOutput & "Plow. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .Hitch Then sOutput = sOutput & "Hitch. "
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .Pin <> "none" Then sOutput = CStr(CDbl(sOutput) + .Pin + CDbl(" pin. "))
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For	Each element In m_oCurrentVeh.Components
				If TypeOf element Is clsWheel Then
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Wheelblades. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If element.Wheelblades <> "none" Then sOutput = CDbl(sOutput) + element.Wheelblades & " wheelblades. "
					'UPGRADE_WARNING: Couldn't resolve default property of object element.snowtires. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If element.snowtires Then sOutput = sOutput & "Snow tires. "
					'UPGRADE_WARNING: Couldn't resolve default property of object element.racingtires. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If element.racingtires Then sOutput = sOutput & "Racing tires. "
					'UPGRADE_WARNING: Couldn't resolve default property of object element.PunctureResistant. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If element.PunctureResistant Then sOutput = sOutput & "Puncture resistant tires. "
				End If
			Next element
		End With
		If sOutput <> "" Then sOutput = "Other Surface Features: " & sOutput
		GetOtherSurfaceFeatures = sOutput
		Exit Function
err_Renamed: 
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Debug.Print("modTextOutput:GetOtherSurfaceFeatures --  Error #" & Err.Number & " " & TypeName(element) & " " & Err.Description)
	End Function
	Private Function GetDefensiveSurfaceFeatures() As String
		Dim m_oCurrentVeh As Object
		Dim element As Object
		Dim sDefensive As String
		Dim sOutput2 As String
		
		On Error GoTo err_Renamed
		'defensive surface features found in Armor classes
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				
				Case ArmorComplexFacing, ArmorBasicFacing, ArmorGunShield, ArmorLocation, ArmorOverall, ArmorWheelGuard
					
					sDefensive = ""
					'UPGRADE_WARNING: Couldn't resolve default property of object element.rap. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If element.rap Then
						If sDefensive <> "" Then
							sDefensive = sDefensive & ", reactive armor"
						Else
							sDefensive = sDefensive & "reactive armor"
						End If
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object element.electrified. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If element.electrified Then
						If sDefensive <> "" Then
							sDefensive = sDefensive & ", electrified"
						Else
							sDefensive = sDefensive & "electrified"
						End If
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object element.thermal. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If element.thermal Then
						If sDefensive <> "" Then
							sDefensive = sDefensive & ", thermal superconductor armor"
						Else
							sDefensive = sDefensive & "thermal superconductor armor"
						End If
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object element.radiation. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If element.radiation Then
						If sDefensive <> "" Then
							sDefensive = sDefensive & ", radiation shielding"
						Else
							sDefensive = sDefensive & "radiation shielding"
						End If
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object element.coating. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If element.coating <> "none" Then
						If sDefensive <> "" Then
							'UPGRADE_WARNING: Couldn't resolve default property of object element.coating. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sDefensive = sDefensive & ", " & element.coating & " coating"
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object element.coating. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sDefensive = sDefensive & element.coating & " coating"
						End If
					End If
					If sDefensive <> "" Then
						'UPGRADE_WARNING: Couldn't resolve default property of object element.LogicalParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sOutput2 = sOutput2 & " On " & m_oCurrentVeh.Components(element.LogicalParent).Description & ": " & sDefensive & "."
					End If
			End Select
		Next element
		
		If sOutput2 <> "" Then sOutput2 = "Defensive Surface Features: " & sOutput2
		GetDefensiveSurfaceFeatures = sOutput2
		Exit Function
err_Renamed: 
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Debug.Print("modTextOutput:GetDefensiveSurfaceFeatures --  Error #" & Err.Number & " " & TypeName(element) & " " & Err.Description)
	End Function
	
	Private Function GetTopDeckSurfaceFeatures() As String
		Dim m_oCurrentVeh As Object
		Dim element As Object
		Dim sTopDeck As String
		Dim sOutput2 As String
		Dim sDeckType As String
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				
				Case Body, Superstructure
					'UPGRADE_WARNING: Couldn't resolve default property of object element.TopDeck. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If element.TopDeck Then
						sTopDeck = ""
						'UPGRADE_WARNING: Couldn't resolve default property of object element.covereddeckarea. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If element.covereddeckarea <> 0 Then
							If sTopDeck <> "" Then
								'UPGRADE_WARNING: Couldn't resolve default property of object element.covereddeckarea. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sTopDeck = sTopDeck & ", " & VB6.Format(element.covereddeckarea, Settings.FormatString) & "sq ft covered"
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object element.covereddeckarea. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sTopDeck = sTopDeck & VB6.Format(element.covereddeckarea, Settings.FormatString) & "sq ft covered"
							End If
						End If
						'UPGRADE_WARNING: Couldn't resolve default property of object element.FlightDeckArea. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If element.FlightDeckArea <> 0 Then
							'get the decktype
							'UPGRADE_WARNING: Couldn't resolve default property of object element.flightdeckoption. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If element.flightdeckoption = "none" Then
								sDeckType = "flight deck"
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object element.flightdeckoption. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sDeckType = element.flightdeckoption
							End If
							If sTopDeck <> "" Then
								'UPGRADE_WARNING: Couldn't resolve default property of object element.flightdecklength. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object element.FlightDeckArea. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sTopDeck = sTopDeck & ", " & VB6.Format(element.FlightDeckArea, Settings.FormatString) & " sq ft " & sDeckType & " with a length of " & VB6.Format(element.flightdecklength, Settings.FormatString) & " ft"
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object element.flightdecklength. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object element.FlightDeckArea. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sTopDeck = sTopDeck & VB6.Format(element.FlightDeckArea, Settings.FormatString) & "sq ft " & sDeckType & " with a length of " & VB6.Format(element.flightdecklength, Settings.FormatString) & " ft"
							End If
						End If
						
						If sTopDeck <> "" Then
							'UPGRADE_WARNING: Couldn't resolve default property of object element.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput2 = sOutput2 & " On " & element.Description & ": " & sTopDeck & "."
						End If
					End If
			End Select
		Next element
		
		If sOutput2 <> "" Then sOutput2 = "Top Deck: " & sOutput2
		GetTopDeckSurfaceFeatures = sOutput2
		Exit Function
err_Renamed: 
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Debug.Print("modTextOutput:GetTopDeckSurfaceFeatures --  Error #" & Err.Number & " " & TypeName(element) & " " & Err.Description)
	End Function
	
	Private Function GetWeaponBaysAndHardpoints() As String
		Dim m_oCurrentVeh As Object
		Dim sOutput As String
		Dim sOutput2 As String
		Dim element As Object
		Dim sngLoad As Single
		Dim sngLoad2 As Single
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				Case WeaponBay
					'UPGRADE_WARNING: Couldn't resolve default property of object element.abbrev. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sOutput = CStr(CDbl(sOutput) + element.abbrev + CDbl(" "))
					'UPGRADE_WARNING: Couldn't resolve default property of object element.loadcapacity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sngLoad = sngLoad + element.loadcapacity
				Case HardPoint
					'UPGRADE_WARNING: Couldn't resolve default property of object element.abbrev. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sOutput2 = CStr(CDbl(sOutput2) + element.abbrev + CDbl(" "))
					'UPGRADE_WARNING: Couldn't resolve default property of object element.loadcapacity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					sngLoad2 = sngLoad2 + element.loadcapacity
			End Select
		Next element
		
		If sOutput <> "" Then
			sOutput = "Weapon bays: " & sOutput & "Total weapon bay load " & VB6.Format(sngLoad, Settings.FormatString) & " lbs."
			
		End If
		If sOutput2 <> "" Then
			sOutput2 = "Hardpoints: " & sOutput2 & "Total hardpoint load " & VB6.Format(sngLoad2, Settings.FormatString) & " lbs."
		End If
		If (sOutput <> "") And (sOutput2 <> "") Then
			sOutput = sOutput & vbNewLine & vbNewLine & sOutput2
		Else
			sOutput = sOutput & sOutput2
		End If
		GetWeaponBaysAndHardpoints = sOutput
		Exit Function
err_Renamed: 
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Debug.Print("modTextOutput:GetWeaponBaysAndHardpoints --  Error #" & Err.Number & " " & TypeName(element) & " " & Err.Description)
	End Function
	
	Private Function GetVisionAndDetailsOutput() As String
		Dim m_oCurrentVeh As Object
		'user entered vision and details
		Dim sOutput As String
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With m_oCurrentVeh.Description
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .Details <> "" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sOutput = CStr(CDbl("Details: ") + .Details)
			End If
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If .Vision <> "" Then
				'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sOutput = CStr(CDbl(sOutput & vbNewLine & "Vision: ") + .Vision)
			End If
		End With
		
		GetVisionAndDetailsOutput = sOutput
		Exit Function
err_Renamed: 
		Debug.Print("modTextOutput:GetVisionDetailsOutput --  Error #" & Err.Number & " " & Err.Description)
	End Function
	Private Function GetPerformanceOutput() As String
		Dim m_oCurrentVeh As Object
		Dim element As Object
		Dim sOutput As String
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.PerformanceProfiles
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If element.Datatype = PERFORMANCEPROFILE Then
				
				With element
					'UPGRADE_WARNING: Couldn't resolve default property of object element.PerformanceType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Select Case .PerformanceType
						Case "Air"
							'UPGRADE_WARNING: Couldn't resolve default property of object element.Key. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = CStr(CDbl(sOutput) + .Key + CDbl(": "))
							'UPGRADE_WARNING: Couldn't resolve default property of object element.aStallSpeed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & "Stall Speed " & VB6.Format(.aStallSpeed, "standard") & " mph."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.aMotiveThrust. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " Aerial motive thrust " & VB6.Format(.aMotiveThrust, "standard") & " lbs."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.aDrag. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " Aerodynamic drag " & .aDrag & "."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.aTopSpeed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " Top speed " & VB6.Format(.aTopSpeed, "standard") & " mph."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.aAcceleration. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " aAccel " & VB6.Format(.aAcceleration, "standard") & " mph/s."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.aManeuverability. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " aMR " & .aManeuverability & "."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.aStability. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " aSR " & .aStability & "."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.aDeceleration. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " aDecel " & VB6.Format(.aDeceleration, "standard") & " mph/s."
							
						Case "Ground"
							'UPGRADE_WARNING: Couldn't resolve default property of object element.Key. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = CStr(CDbl(sOutput) + .Key + CDbl(": "))
							'UPGRADE_WARNING: Couldn't resolve default property of object element.gTopSpeed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " Speed " & VB6.Format(.gTopSpeed, "standard") & " mph."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.gAcceleration. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " gAccel " & VB6.Format(.gAcceleration, "standard") & " mph/s."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.gDeceleration. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " gDecel " & VB6.Format(.gDeceleration, "standard") & " mph/s."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.gStability. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " gSR " & .gStability & "."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.gManeuverability. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " gMR " & .gManeuverability & "."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.gPressureDescription. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " " & VB6.Format(.gPressureDescription, "standard") & " ground pressure."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.gOffRoad. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " Off road speed " & VB6.Format(.gOffRoad, "standard") & " mph/s."
							
							
						Case "Hovercraft"
							'UPGRADE_WARNING: Couldn't resolve default property of object element.Key. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = CStr(CDbl(sOutput) + .Key + CDbl(": "))
							'UPGRADE_WARNING: Couldn't resolve default property of object element.hHoverAltitude. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " Hover Altitude " & .hHoverAltitude & " feet."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.hMotiveThrust. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " Thrust " & VB6.Format(.hMotiveThrust, "standard") & " lbs."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.hTopSpeed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " Speed " & VB6.Format(.hTopSpeed, "standard") & " mph."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.hDrag. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " Drag " & .hDrag & "."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.hAcceleration. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " hAccel " & VB6.Format(.hAcceleration, "standard") & " mph/s."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.hstability. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " hSR " & .hstability & "."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.hmaneuverability. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " hMR " & .hmaneuverability & " g."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.hDeceleration. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " hDecel " & VB6.Format(.hDeceleration, "standard") & " mph/s."
							
							
						Case "Mag-Lev"
							'UPGRADE_WARNING: Couldn't resolve default property of object element.Key. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = CStr(CDbl(sOutput) + .Key + CDbl(": "))
							'UPGRADE_WARNING: Couldn't resolve default property of object element.mlMotiveThrust. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " Thrust " & VB6.Format(.mlMotiveThrust, "standard") & " lbs."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.mlTopSpeed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " Speed " & VB6.Format(.mlTopSpeed, "standard") & " mph."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.mlStallSpeed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " Stall Speed " & VB6.Format(.mlStallSpeed, "standard") & " mph."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.mlDrag. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " mDrag " & .mlDrag & "."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.mlAcceleration. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " mAccel " & VB6.Format(.mlAcceleration, "standard") & " mph/s."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.mlStability. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " mSR " & .mlStability & "."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.mlManeuverability. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " mMR " & .mlManeuverability & "."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.mlDeceleration. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " mDecel " & VB6.Format(.mlDeceleration, "standard") & " mph/s."
							
						Case "Water"
							'UPGRADE_WARNING: Couldn't resolve default property of object element.Key. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = CStr(CDbl(sOutput) + .Key + CDbl(": "))
							'UPGRADE_WARNING: Couldn't resolve default property of object element.wHydroDrag. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " Hydrodynamic drag " & VB6.Format(.wHydroDrag, "standard") & "."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.wTotalAquaticThrust. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " Aquatic motive thrust " & VB6.Format(.wTotalAquaticThrust, "standard") & " lbs."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.wTopSpeed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " Speed " & VB6.Format(.wTopSpeed, "standard") & " mph."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.wHydrofoilSpeed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " Hydrofoil Speed " & VB6.Format(.wHydrofoilSpeed, "standard") & " mph."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.wPlaningSpeed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " Planing Speed " & VB6.Format(.wPlaningSpeed, "standard") & " mph."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.wAcceleration. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " wAccel " & VB6.Format(.wAcceleration, "standard") & " mph/s."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.wManeuverability. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " wMR " & .wManeuverability & "."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.wStability. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " wSR  " & .wStability & "."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.wDeceleration. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " wDecel " & VB6.Format(.wDeceleration, "standard") & " mph/s."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.wIDeceleration. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If .wIDeceleration > 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object element.wIDeceleration. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sOutput = sOutput & " Incr wDecel " & VB6.Format(.wIDeceleration, "standard") & " mph/s."
							End If
							'UPGRADE_WARNING: Couldn't resolve default property of object element.wDraft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " wDraft " & VB6.Format(.wDraft, "standard") & " feet."
							
						Case "Submerged"
							'UPGRADE_WARNING: Couldn't resolve default property of object element.Key. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = CStr(CDbl(sOutput) + .Key + CDbl(": "))
							'UPGRADE_WARNING: Couldn't resolve default property of object element.suTotalAquaticThrust. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & "suThrust " & VB6.Format(.suTotalAquaticThrust, "standard") & " lbs."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.suHydroDrag. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " suDrag " & .suHydroDrag & "."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.suTopSpeed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " suSpeed " & VB6.Format(.suTopSpeed, "standard") & " mph."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.suAcceleration. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " suAccel " & VB6.Format(.suAcceleration, "standard") & " mph/s."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.suDeceleration. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " suDecel " & VB6.Format(.suDeceleration, "standard") & " mph/s."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.suIDeceleration. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If .suIDeceleration > 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object element.suIDeceleration. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sOutput = sOutput & " Incr suDecel " & VB6.Format(.suIDeceleration, "standard") & " mph/s."
							End If
							'UPGRADE_WARNING: Couldn't resolve default property of object element.suStability. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " suSR " & .suStability & "."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.suManeuverability. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " suMR " & .suManeuverability & "."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.suDraft. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " Draft " & VB6.Format(.suDraft, "standard") & " feet."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.suCrushDepth. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If .suCrushDepth = -1 Then
								sOutput = sOutput & "No Crush Depth"
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object element.suCrushDepth. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								sOutput = sOutput & " Crush Depth " & VB6.Format(.suCrushDepth, "standard") & " yards."
							End If
							
						Case "Space"
							'UPGRADE_WARNING: Couldn't resolve default property of object element.Key. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = CStr(CDbl(sOutput) + .Key + CDbl(": "))
							'UPGRADE_WARNING: Couldn't resolve default property of object element.sMotiveThrust. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " Thrust " & VB6.Format(.sMotiveThrust, "standard") & " lbs."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.sAccelerationG. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " sAccel " & VB6.Format(.sAccelerationG, "standard") & " g."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.sAccelerationMPH. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " sAccel " & VB6.Format(.sAccelerationMPH, "standard") & " mph/s."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.sTurnAroundTime. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " Turn Around " & VB6.Format(.sTurnAroundTime, "standard") & " secs."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.sManeuverability. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " sMR " & VB6.Format(.sManeuverability, "standard") & "."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.sHyperSpeed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " Hyper " & VB6.Format(.sHyperSpeed, "standard") & " parsecs per day."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.sWarpSpeed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							sOutput = sOutput & " Warp " & VB6.Format(.sWarpSpeed, "standard") & " parsecs per day."
							'UPGRADE_WARNING: Couldn't resolve default property of object element.sJumpDriveable. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If .sJumpDriveable Then
								sOutput = sOutput & " Has jump capabilities."
							End If
							'UPGRADE_WARNING: Couldn't resolve default property of object element.sTeleportationDriveable. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If .sTeleportationDriveable Then
								sOutput = sOutput & " Has teleportation drive capabilities."
							End If
					End Select
					sOutput = sOutput & vbNewLine & vbNewLine
				End With
			End If
		Next element
		
		GetPerformanceOutput = sOutput
		Exit Function
err_Renamed: 
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Debug.Print("modTextOutput:GetPerformanceOutput --  Error #" & Err.Number & " " & TypeName(element) & " " & Err.Description)
	End Function
	
	Private Function GetDetailedWeaponStats() As String
		Dim m_oCurrentVeh As Object
		Dim element As Object
		Dim sOutput As String
		Dim aHeader, bHeader As Object
		Dim cHeader As Boolean
		Dim gun() As Object
		Dim j, i, k As Object
		Dim l As Integer
		Dim iLength As Integer
		Dim iOldLength As Integer
		Dim iPropID As Integer
		
		
		On Error GoTo err_Renamed
		
		'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		i = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		j = 1
		'UPGRADE_WARNING: Couldn't resolve default property of object k. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		k = 1
		l = 1
		'//guns and artillery
		'UPGRADE_WARNING: Couldn't resolve default property of object bHeader. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		bHeader = False
		'UPGRADE_WARNING: Lower bound of array gun was changed from 1,0 to 0,0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim gun(15, 1)
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				Case StoneThrower, BoltThrower, RepeatingBoltThrower, MuzzleLoader, BreechLoader, ManualRepeater, Revolver, MechanicalGatling, SlowAutoloader, FastAutoloader, lightAutomatic, HeavyAutomatic, ElectricGatling
					
					
					'UPGRADE_WARNING: Couldn't resolve default property of object bHeader. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If bHeader = False Then
						'we havent printed our header yet so do it now
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(1, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(1, 1) = "Name"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(2, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(2, 1) = "Malf"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(3, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(3, 1) = "Type"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(4, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(4, 1) = "Damage"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(5, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(5, 1) = "SS"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(6, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(6, 1) = "Acc"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(7, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(7, 1) = "1/2D"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(8, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(8, 1) = "Max"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(9, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(9, 1) = "RoF"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(10, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(10, 1) = "Weight"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(11, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(11, 1) = "Cost"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(12, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(12, 1) = "WPS"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(13, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(13, 1) = "VPS"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(14, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(14, 1) = "CPS"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(15, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(15, 1) = "Ldrs."
						'UPGRADE_WARNING: Couldn't resolve default property of object bHeader. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						bHeader = True
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					i = i + 1
					'UPGRADE_WARNING: Lower bound of array gun was changed from 1,0 to 0,0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
					ReDim Preserve gun(15, i)
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.CustomDescription. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(1, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(1, i) = element.CustomDescription
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Malfunction. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(2, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(2, i) = element.Malfunction
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.TypeDamage1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(3, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(3, i) = element.TypeDamage1
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Damage1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(4, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(4, i) = element.Damage1
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.SnapShot. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(5, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(5, i) = element.SnapShot
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Accuracy. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(6, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(6, i) = element.Accuracy
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.halfDamage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(7, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(7, i) = element.halfDamage
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.MaxRange. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(8, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(8, i) = element.MaxRange
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.sRoF. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(9, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(9, i) = element.sRoF
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Weight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(10, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(10, i) = element.Weight
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Cost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(11, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(11, i) = element.Cost
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.WPS. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(12, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(12, i) = element.WPS
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.VPS. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(13, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(13, i) = element.VPS
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.CPS. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(14, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(14, i) = element.CPS
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Loaders. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(15, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(15, i) = element.Loaders
			End Select
		Next element
		'//now we must pad each row item with spaces so that they are all the same length
		'UPGRADE_WARNING: Couldn't resolve default property of object gun(1, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If gun(1, 1) <> "" Then
			For iPropID = 1 To 15
				iLength = 0
				iOldLength = 0
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				For j = 1 To i
					'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Maximum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					iLength = modPerformance.Maximum(iLength, Len(gun(iPropID, j)))
				Next 
				iLength = iLength + 1 '//we need 1 space seperation
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				For j = 1 To i
					'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					iOldLength = Len(gun(iPropID, j))
					For k = 1 To iLength - iOldLength
						'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(iPropID, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(iPropID, j) = gun(iPropID, j) & " "
					Next 
				Next 
			Next 
			'//finally we can output it all
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For j = 1 To i
				sOutput = sOutput & sLineBreak
				'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(15, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(14, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(13, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(12, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(11, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(10, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(9, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(8, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(7, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(6, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(5, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(4, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(3, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(2, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sOutput = sOutput & gun(1, j) & gun(2, j) & gun(3, j) & gun(4, j) & gun(5, j) & gun(6, j) & gun(7, j) & gun(8, j) & gun(9, j) & gun(10, j) & gun(11, j) & gun(12, j) & gun(13, j) & gun(14, j) & gun(15, j)
			Next 
			sOutput = sOutput & sLineBreak
			'UPGRADE_WARNING: Couldn't resolve default property of object gun(1, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			gun(1, 1) = ""
		End If
		'////////////////////////////////////////////////////////////////////
		'//Beam Weapons
		'UPGRADE_WARNING: Couldn't resolve default property of object bHeader. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		bHeader = False
		'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		i = 1
		'UPGRADE_WARNING: Lower bound of array gun was changed from 1,0 to 0,0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim gun(12, 1)
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				Case BlueGreenLaser, RainbowLaser, Laser, UVLaser, IRLaser, Disruptor, ChargedParticleBeam, NeutralParticleBeam, Flamer, Screamer, Stunner, ParalysisBeam, XRayLaser, FusionBeam, GravityBeam, AntiparticleBeam, Graser, Disintegrator, Displacer, BeamedPowerTransmitter, MilitaryParalysisBeam, EnergyDrill
					
					'UPGRADE_WARNING: Couldn't resolve default property of object bHeader. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If bHeader = False Then
						'we havent printed our header yet so do it now
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(1, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(1, 1) = "Name"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(2, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(2, 1) = "Malf"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(3, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(3, 1) = "Type"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(4, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(4, 1) = "Damage"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(5, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(5, 1) = "SS"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(6, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(6, 1) = "Acc"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(7, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(7, 1) = "1/2D"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(8, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(8, 1) = "Max"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(9, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(9, 1) = "RoF"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(10, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(10, 1) = "Weight"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(11, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(11, 1) = "Cost"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(12, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(12, 1) = "Power"
						'UPGRADE_WARNING: Couldn't resolve default property of object bHeader. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						bHeader = True
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					i = i + 1
					'UPGRADE_WARNING: Lower bound of array gun was changed from 1,0 to 0,0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
					ReDim Preserve gun(12, i)
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.CustomDescription. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(1, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(1, i) = element.CustomDescription
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Malfunction. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(2, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(2, i) = element.Malfunction
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.TypeDamage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(3, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(3, i) = element.TypeDamage
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Damage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(4, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(4, i) = element.Damage
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.SnapShot. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(5, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(5, i) = element.SnapShot
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Accuracy. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(6, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(6, i) = element.Accuracy
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.halfDamage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(7, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(7, i) = element.halfDamage
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.MaxRange. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(8, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(8, i) = element.MaxRange
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.rof. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(9, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(9, i) = element.rof
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Weight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(10, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(10, i) = element.Weight
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Cost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(11, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(11, i) = element.Cost
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.PowerReqt. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(12, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(12, i) = element.PowerReqt
			End Select
		Next element
		'//now we must pad each row item with spaces so that they are all the same length
		'UPGRADE_WARNING: Couldn't resolve default property of object gun(1, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If gun(1, 1) <> "" Then
			For iPropID = 1 To 12
				iLength = 0
				iOldLength = 0
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				For j = 1 To i
					'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Maximum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					iLength = modPerformance.Maximum(iLength, Len(gun(iPropID, j)))
				Next 
				iLength = iLength + 1 '//we need 1 space seperation
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				For j = 1 To i
					'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					iOldLength = Len(gun(iPropID, j))
					For k = 1 To iLength - iOldLength
						'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(iPropID, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(iPropID, j) = gun(iPropID, j) & " "
					Next 
				Next 
			Next 
			'//finally we can output it all
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For j = 1 To i
				sOutput = sOutput & sLineBreak
				'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(12, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(11, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(10, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(9, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(8, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(7, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(6, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(5, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(4, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(3, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(2, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sOutput = sOutput & gun(1, j) & gun(2, j) & gun(3, j) & gun(4, j) & gun(5, j) & gun(6, j) & gun(7, j) & gun(8, j) & gun(9, j) & gun(10, j) & gun(11, j) & gun(12, j)
			Next 
			sOutput = sOutput & sLineBreak
			'UPGRADE_WARNING: Couldn't resolve default property of object gun(1, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			gun(1, 1) = ""
		End If
		'////////////////////////////////////////////////////////////////
		'//Bombs, missiles and torps
		'UPGRADE_WARNING: Couldn't resolve default property of object bHeader. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		bHeader = False
		'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		i = 1
		'UPGRADE_WARNING: Lower bound of array gun was changed from 1,0 to 0,0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim gun(13, 1)
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				Case IronBomb, RetardedBomb, SmartBomb, SelfDestructSystem, ContactMine, ProximityMine, PressureTriggerMine, CommandTriggerMine, SmartTriggerMine, UnGuidedMissile, UnGuidedTorpedo, GuidedMissile, GuidedTorpedo
					
					'UPGRADE_WARNING: Couldn't resolve default property of object bHeader. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If bHeader = False Then
						'we havent printed our header yet so do it now
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(1, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(1, 1) = "Name"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(2, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(2, 1) = "Malf"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(3, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(3, 1) = "Guid"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(4, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(4, 1) = "Type"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(5, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(5, 1) = "Damage"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(6, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(6, 1) = "Spd"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(7, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(7, 1) = "End"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(8, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(8, 1) = "Max"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(9, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(9, 1) = "Min"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(10, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(10, 1) = "Skill"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(11, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(11, 1) = "WPS"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(12, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(12, 1) = "VPS"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(13, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(13, 1) = "CPS"
						
						'UPGRADE_WARNING: Couldn't resolve default property of object bHeader. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						bHeader = True
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					i = i + 1
					'UPGRADE_WARNING: Lower bound of array gun was changed from 1,0 to 0,0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
					ReDim Preserve gun(13, i)
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.CustomDescription. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(1, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(1, i) = element.CustomDescription
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Malfunction. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(2, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(2, i) = element.Malfunction
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.GuidanceSystem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(3, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(3, i) = element.GuidanceSystem
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.TypeDamage1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(4, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(4, i) = element.TypeDamage1
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Damage1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(5, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(5, i) = element.Damage1
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Speed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(6, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(6, i) = element.Speed
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Endurance. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(7, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(7, i) = element.Endurance
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.MaxRange. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(8, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(8, i) = element.MaxRange
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.MinRange. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(9, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(9, i) = element.MinRange
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Skill. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(10, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(10, i) = element.Skill
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Quantity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Weight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(11, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(11, i) = element.Weight / element.Quantity
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Volume. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(12, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(12, i) = element.Volume
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Quantity. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Cost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(13, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(13, i) = element.Cost / element.Quantity
			End Select
		Next element
		'//now we must pad each row item with spaces so that they are all the same length
		'UPGRADE_WARNING: Couldn't resolve default property of object gun(1, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If gun(1, 1) <> "" Then
			For iPropID = 1 To 13
				iLength = 0
				iOldLength = 0
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				For j = 1 To i
					'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Maximum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					iLength = modPerformance.Maximum(iLength, Len(gun(iPropID, j)))
				Next 
				iLength = iLength + 1 '//we need 1 space seperation
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				For j = 1 To i
					'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					iOldLength = Len(gun(iPropID, j))
					For k = 1 To iLength - iOldLength
						'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(iPropID, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(iPropID, j) = gun(iPropID, j) & " "
					Next 
				Next 
			Next 
			'//finally we can output it all
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For j = 1 To i
				sOutput = sOutput & sLineBreak
				'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(13, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(12, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(11, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(10, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(9, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(8, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(7, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(6, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(5, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(4, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(3, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(2, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sOutput = sOutput & gun(1, j) & gun(2, j) & gun(3, j) & gun(4, j) & gun(5, j) & gun(6, j) & gun(7, j) & gun(8, j) & gun(9, j) & gun(10, j) & gun(11, j) & gun(12, j) & gun(13, j)
			Next 
			sOutput = sOutput & sLineBreak
			'UPGRADE_WARNING: Couldn't resolve default property of object gun(1, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			gun(1, 1) = ""
		End If
		'/////////////////////////////////////////////////////////////
		'//Liquid projectors
		'UPGRADE_WARNING: Couldn't resolve default property of object bHeader. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		bHeader = False
		'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		i = 1
		'UPGRADE_WARNING: Lower bound of array gun was changed from 1,0 to 0,0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim gun(13, 1)
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				Case FlameThrower, WaterCannon
					
					'UPGRADE_WARNING: Couldn't resolve default property of object bHeader. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If bHeader = False Then
						'we havent printed our header yet so do it now
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(1, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(1, 1) = "Name"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(2, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(2, 1) = "Malf"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(3, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(3, 1) = "Type"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(4, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(4, 1) = "Damage"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(5, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(5, 1) = "SS"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(6, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(6, 1) = "Acc"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(7, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(7, 1) = "1/2D"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(8, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(8, 1) = "Max"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(9, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(9, 1) = "RoF"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(10, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(10, 1) = "Weight"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(11, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(11, 1) = "Cost"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(12, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(12, 1) = "WPS"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(13, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(13, 1) = "CPS"
						'UPGRADE_WARNING: Couldn't resolve default property of object bHeader. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						bHeader = True
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					i = i + 1
					'UPGRADE_WARNING: Lower bound of array gun was changed from 1,0 to 0,0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
					ReDim Preserve gun(13, i)
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.CustomDescription. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(1, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(1, i) = element.CustomDescription
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Malfunction. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(2, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(2, i) = element.Malfunction
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.TypeDamage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(3, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(3, i) = element.TypeDamage
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Damage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(4, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(4, i) = element.Damage
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.SnapShot. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(5, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(5, i) = element.SnapShot
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Accuracy. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(6, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(6, i) = element.Accuracy
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.halfDamage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(7, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(7, i) = element.halfDamage
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.MaxRange. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(8, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(8, i) = element.MaxRange
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.rof. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(9, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(9, i) = element.rof
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Weight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(10, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(10, i) = element.Weight
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Cost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(11, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(11, i) = element.Cost
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.WPS. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(12, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(12, i) = element.WPS
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.CPS. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(13, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(13, i) = element.CPS
			End Select
		Next element
		'//now we must pad each row item with spaces so that they are all the same length
		'UPGRADE_WARNING: Couldn't resolve default property of object gun(1, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If gun(1, 1) <> "" Then
			For iPropID = 1 To 13
				iLength = 0
				iOldLength = 0
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				For j = 1 To i
					'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Maximum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					iLength = modPerformance.Maximum(iLength, Len(gun(iPropID, j)))
				Next 
				iLength = iLength + 1 '//we need 1 space seperation
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				For j = 1 To i
					'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					iOldLength = Len(gun(iPropID, j))
					For k = 1 To iLength - iOldLength
						'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(iPropID, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(iPropID, j) = gun(iPropID, j) & " "
					Next 
				Next 
			Next 
			'//finally we can output it all
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For j = 1 To i
				sOutput = sOutput & sLineBreak
				'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(13, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(12, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(11, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(10, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(9, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(8, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(7, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(6, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(5, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(4, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(3, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(2, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sOutput = sOutput & gun(1, j) & gun(2, j) & gun(3, j) & gun(4, j) & gun(5, j) & gun(6, j) & gun(7, j) & gun(8, j) & gun(9, j) & gun(10, j) & gun(11, j) & gun(12, j) & gun(13, j)
			Next 
			sOutput = sOutput & sLineBreak
			'UPGRADE_WARNING: Couldn't resolve default property of object gun(1, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			gun(1, 1) = ""
		End If
		'///////////////////////////////////////////////////////////////
		'//Launchers
		'UPGRADE_WARNING: Couldn't resolve default property of object bHeader. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		bHeader = False
		'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		i = 1
		'UPGRADE_WARNING: Lower bound of array gun was changed from 1,0 to 0,0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim gun(7, 1)
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		For	Each element In m_oCurrentVeh.Components
			'UPGRADE_WARNING: Couldn't resolve default property of object element.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case element.Datatype
				Case DisposableLauncher, MuzzleloadingLauncher, BreechloadingLauncher, ManualRepeaterLauncher, SlowAutoLoaderLauncher, FastAutoLoaderLauncher, RevolverLauncher, lightAutomaticLauncher, HeavyAutomaticLauncher
					
					'UPGRADE_WARNING: Couldn't resolve default property of object bHeader. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If bHeader = False Then
						'we havent printed our header yet so do it now
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(1, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(1, 1) = "Name"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(2, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(2, 1) = "SS"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(3, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(3, 1) = "RoF"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(4, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(4, 1) = "Weight"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(5, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(5, 1) = "Cost"
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(6, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(6, 1) = "Ldrs."
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(7, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(7, 1) = "Rating"
						'UPGRADE_WARNING: Couldn't resolve default property of object bHeader. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						bHeader = True
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					i = i + 1
					'UPGRADE_WARNING: Lower bound of array gun was changed from 1,0 to 0,0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
					ReDim Preserve gun(7, i)
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.CustomDescription. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(1, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(1, i) = element.CustomDescription
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.SnapShot. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(2, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(2, i) = element.SnapShot
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.rof. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(3, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(3, i) = element.rof
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Weight. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(4, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(4, i) = element.Weight
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Cost. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(5, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(5, i) = element.Cost
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.Loaders. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(6, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(6, i) = element.Loaders
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object element.MaxLoad. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gun(7, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gun(7, i) = element.MaxLoad
			End Select
		Next element
		'//now we must pad each row item with spaces so that they are all the same length
		'UPGRADE_WARNING: Couldn't resolve default property of object gun(1, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If gun(1, 1) <> "" Then
			For iPropID = 1 To 7
				iLength = 0
				iOldLength = 0
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				For j = 1 To i
					'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object Maximum(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					iLength = modPerformance.Maximum(iLength, Len(gun(iPropID, j)))
				Next 
				iLength = iLength + 1 '//we need 1 space seperation
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				For j = 1 To i
					'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					iOldLength = Len(gun(iPropID, j))
					For k = 1 To iLength - iOldLength
						'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object gun(iPropID, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gun(iPropID, j) = gun(iPropID, j) & " "
					Next 
				Next 
			Next 
			'//finally we can output it all
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For j = 1 To i
				sOutput = sOutput & sLineBreak
				'UPGRADE_WARNING: Couldn't resolve default property of object j. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(7, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(6, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(5, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(4, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(3, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(2, j). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gun(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				sOutput = sOutput & gun(1, j) & gun(2, j) & gun(3, j) & gun(4, j) & gun(5, j) & gun(6, j) & gun(7, j)
			Next 
			sOutput = sOutput & sLineBreak
			'UPGRADE_WARNING: Couldn't resolve default property of object gun(1, 1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			gun(1, 1) = ""
		End If
		'//add a new line and send then return the entire output value
		sOutput = sOutput & vbNewLine
		GetDetailedWeaponStats = sOutput
		Exit Function
err_Renamed: 
		'UPGRADE_WARNING: TypeName has a new behavior. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		Debug.Print("modTextOutput:GetDetailedWeaponStats --  Error #" & Err.Number & " " & TypeName(element) & " " & Err.Description)
	End Function
	
	Private Function NumericToString(ByVal nNumber As Object) As String
		'//this function accepts a number and if that number is between
		'1 and 10 it will convert them to "One" and "Ten" for instance. If
		'the number is greater than 10 it will just return the number formatted
		'as a string
		Dim retval As String
		
		'NOTE: This is currently only set up to handle longs and not decimals
		'UPGRADE_WARNING: Couldn't resolve default property of object nNumber. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If nNumber >= 1 And nNumber <= 10 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object nNumber. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If nNumber = 1 Then
				retval = "One"
				'retval = "" 'if its a 1 we'll jsut leave it blank since its assumed to be 1 unless noted
				'UPGRADE_WARNING: Couldn't resolve default property of object nNumber. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf nNumber = 2 Then 
				retval = "Two"
				'UPGRADE_WARNING: Couldn't resolve default property of object nNumber. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf nNumber = 3 Then 
				retval = "Three"
				'UPGRADE_WARNING: Couldn't resolve default property of object nNumber. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf nNumber = 4 Then 
				retval = "Four"
				'UPGRADE_WARNING: Couldn't resolve default property of object nNumber. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf nNumber = 5 Then 
				retval = "Five"
				'UPGRADE_WARNING: Couldn't resolve default property of object nNumber. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf nNumber = 6 Then 
				retval = "Six"
				'UPGRADE_WARNING: Couldn't resolve default property of object nNumber. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf nNumber = 7 Then 
				retval = "Seven"
				'UPGRADE_WARNING: Couldn't resolve default property of object nNumber. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf nNumber = 8 Then 
				retval = "Eight"
				'UPGRADE_WARNING: Couldn't resolve default property of object nNumber. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf nNumber = 9 Then 
				retval = "Nine"
				'UPGRADE_WARNING: Couldn't resolve default property of object nNumber. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ElseIf nNumber = 10 Then 
				retval = "Ten"
			End If
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object nNumber. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			retval = "(" & VB6.Format(nNumber) & ")"
		End If
		
		NumericToString = retval
		
	End Function
	
	Private Function RemoveParenthetical(ByVal strIn As String) As Object
		'JAW 2000.060.26
		Dim varSplit As Object
		Dim i As Short
		Dim strTemp As String
		Dim strLocation As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object varSplit. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		varSplit = Split(strIn, "(")
		For i = 1 To UBound(varSplit)
			'UPGRADE_WARNING: Couldn't resolve default property of object varSplit(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			strTemp = varSplit(i - 1)
			If InStr(1, strTemp, ")") Then
				'                If Left(strTemp, 1) = "$" Then
				'                strLocation = ""
				'            Else
				'                strLocation = Left(strTemp, 2)
				'            End If
				'        varSplit(i - 1) = "[" & strLocation & "]" & Split(strTemp, ")")(1)
				'UPGRADE_WARNING: Couldn't resolve default property of object varSplit(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				varSplit(i - 1) = Split(strTemp, ")")(1)
			End If
		Next i
		For i = 1 To UBound(varSplit)
			'UPGRADE_WARNING: Couldn't resolve default property of object varSplit(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object RemoveParenthetical. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			RemoveParenthetical = RemoveParenthetical & varSplit(i - 1)
		Next i
		
	End Function
End Module