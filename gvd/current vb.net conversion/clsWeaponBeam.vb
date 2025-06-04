Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsWeaponBeam_NET.clsWeaponBeam")> Public Class clsWeaponBeam
	
	Private mvarHitPoints As Double
	Private mvarDR As Integer
	Private mvarCustom As Boolean
	Private mvarCost As Double
	Private mvarDatatype As Short
	Private mvarDescription As String
	Private mvarImage As Short
	Private mvarSelectedImage As Short
	Private mvarKey As String
	Private mvarParent As String
	Private mvarQuantity As Short
	Private mvarSurfaceArea As Double
	Private mvarTL As Short
	Private mvarVolume As Double
	Private mvarWeight As Double
	Private mvarRuggedized As Boolean
	Private mvarCustomDescription As String
	Private mvarQuality As String
	Private mvarBeamOutput As Single
	Private mvarCyclicRate As Single
	Private mvarRange As String
	Private mvarCompact As Boolean
	Private mvarDamage As String
	Private mvarKEDamage As Double
	Private mvarTypeDamage As String
	Private mvarhalfDamage As Double
	Private mvarMaxRange As Double
	Private mvarMaxRange2 As Double
	Private mvarVacuumMaxRange2 As Double
	Private mvarVacuumHalfDamage As Double
	Private mvarVacuumMaxRange As Double
	Private mvarAccuracy As Integer
	Private mvarSnapShot As Integer
	Private mvarShots As String
	Private mvarRoF As String
	Private mvarPowerReqt As Double
	Private mvarMount As String
	Private mvarReliable As Boolean
	Private mvarMalfunction As String
	Private mvarFTL As Boolean
	Private mvarPowerCellType As String
	Private mvarPowerCellQuantity As Integer
	Private mvarPowerCellWeight As Single
	Private mvarDirection As String
	Private mvarLocation As String
	Private mvarComment As String
	Private mvarCName As String
	Private mvarEnergyDrill As Boolean
	Private mvarPrintOutput As String
	Private mvarZZInit As Byte
	Private mvarLogicalParent As String
	
	
	Public Property LogicalParent() As String
		Get
			LogicalParent = mvarLogicalParent
		End Get
		Set(ByVal Value As String)
			mvarLogicalParent = Value
		End Set
	End Property
	
	
	
	Public Property PrintOutput() As String
		Get


			PrintOutput = mvarPrintOutput
		End Get
		Set(ByVal Value As String)


			mvarPrintOutput = Value
		End Set
	End Property
	
	
	
	
	Public Property CName() As String
		Get


            CName = mvarCName
        End Get
        Set(ByVal Value As String)


            mvarCName = Value
        End Set
    End Property





    Public Property Comment() As String
        Get


            Comment = mvarComment
        End Get
        Set(ByVal Value As String)


            mvarComment = Value
        End Set
    End Property
	
	
	
	
	
	Public Property Location() As String
		Get


            Location = mvarLocation
        End Get
        Set(ByVal Value As String)


            mvarLocation = Value
        End Set
    End Property




    Public Property Direction() As String
        Get


            Direction = mvarDirection
        End Get
        Set(ByVal Value As String)


            mvarDirection = Value
        End Set
    End Property
	
	
	Public Property PowerCellType() As String
		Get


            PowerCellType = mvarPowerCellType
        End Get
        Set(ByVal Value As String)


            mvarPowerCellType = Value
        End Set
    End Property



    Public Property PowerCellQuantity() As Integer
        Get


            PowerCellQuantity = mvarPowerCellQuantity
        End Get
        Set(ByVal Value As Integer)


            mvarPowerCellQuantity = Value
        End Set
    End Property
	
	
	
	
	Public Property PowerCellWeight() As Double
		Get


            PowerCellWeight = mvarPowerCellWeight
        End Get
        Set(ByVal Value As Double)


            mvarPowerCellWeight = Value
        End Set
    End Property



    Public Property FTL() As Boolean
        Get


            FTL = mvarFTL
        End Get
        Set(ByVal Value As Boolean)


            mvarFTL = Value
        End Set
    End Property
	
	
	Public Property Reliable() As Boolean
		Get


            Reliable = mvarReliable
        End Get
        Set(ByVal Value As Boolean)


            mvarReliable = Value
        End Set
    End Property


    Public Property Malfunction() As String
        Get


            Malfunction = mvarMalfunction
        End Get
        Set(ByVal Value As String)


            mvarMalfunction = Value
        End Set
    End Property
	
	
	
	Public Property Mount() As String
		Get


            Mount = mvarMount
        End Get
        Set(ByVal Value As String)


            mvarMount = Value
        End Set
    End Property



    Public Property PowerReqt() As Double
        Get


            PowerReqt = mvarPowerReqt
        End Get
        Set(ByVal Value As Double)


            mvarPowerReqt = Value
        End Set
    End Property
	
	
	
	
	
	Public Property Shots() As String
		Get


            Shots = mvarShots
        End Get
        Set(ByVal Value As String)


            mvarShots = Value
        End Set
    End Property





    Public Property SnapShot() As Integer
        Get


            SnapShot = mvarSnapShot
        End Get
        Set(ByVal Value As Integer)


            mvarSnapShot = Value
        End Set
    End Property
	
	
	
	
	
	Public Property Accuracy() As Integer
		Get


            Accuracy = mvarAccuracy
        End Get
        Set(ByVal Value As Integer)


            mvarAccuracy = Value
        End Set
    End Property





    Public Property MaxRange() As Double
        Get


            MaxRange = mvarMaxRange
        End Get
        Set(ByVal Value As Double)


            mvarMaxRange = Value
        End Set
    End Property
	
	
	
	Public Property MaxRange2() As Single
		Get


            MaxRange2 = mvarMaxRange2
        End Get
        Set(ByVal Value As Single)


            mvarMaxRange2 = Value
        End Set
    End Property


    Public Property VacuumMaxRange() As Double
        Get


            VacuumMaxRange = mvarVacuumMaxRange
        End Get
        Set(ByVal Value As Double)


            mvarVacuumMaxRange = Value
        End Set
    End Property
	
	
	Public Property VacuumMaxRange2() As Single
		Get


            VacuumMaxRange2 = mvarVacuumMaxRange2
        End Get
        Set(ByVal Value As Single)


            mvarVacuumMaxRange2 = Value
        End Set
    End Property



    Public Property VacuumHalfDamage() As Single
        Get


            VacuumHalfDamage = mvarVacuumHalfDamage
        End Get
        Set(ByVal Value As Single)


            mvarVacuumHalfDamage = Value
        End Set
    End Property
	
	
	
	
	Public Property halfDamage() As Double
		Get


            halfDamage = mvarhalfDamage
        End Get
        Set(ByVal Value As Double)


            mvarhalfDamage = Value
        End Set
    End Property





    Public Property TypeDamage() As String
        Get


            TypeDamage = mvarTypeDamage
        End Get
        Set(ByVal Value As String)


            mvarTypeDamage = Value
        End Set
    End Property
	
	
	
	Public Property KEDamage() As Double
		Get


            KEDamage = mvarKEDamage
        End Get
        Set(ByVal Value As Double)


            mvarKEDamage = Value
        End Set
    End Property



    Public Property Damage() As String
        Get


            Damage = mvarDamage
        End Get
        Set(ByVal Value As String)


            mvarDamage = Value
        End Set
    End Property
	
	
	
	
	
	Public Property Compact() As Boolean
		Get


            Compact = mvarCompact
        End Get
        Set(ByVal Value As Boolean)


            mvarCompact = Value
        End Set
    End Property



    Public Property EnergyDrill() As Boolean
        Get


            EnergyDrill = mvarEnergyDrill
        End Get
        Set(ByVal Value As Boolean)


            mvarEnergyDrill = Value
        End Set
    End Property
	
	
	
	Public Property Range() As String
		Get


            Range = mvarRange
        End Get
        Set(ByVal Value As String)


            mvarRange = Value
        End Set
    End Property



    Public Property rof() As String
        Get


            rof = mvarRoF
        End Get
        Set(ByVal Value As String)



            mvarRoF = Value
            If mvarZZInit = 0 Then Exit Property
            On Error Resume Next
            Dim num As Single
            Dim arr() As String

            arr = Split(mvarRoF, "/")

            If UBound(arr) = 0 Then
                num = Val(mvarRoF)
            Else
                num = CDbl(arr(0)) / CDbl(arr(1))
            End If

            CyclicRate = num
        End Set
    End Property
	
	
	
	Public Property CyclicRate() As Single
		Get


            CyclicRate = mvarCyclicRate
        End Get
        Set(ByVal Value As Single)


            mvarCyclicRate = Value
            If mvarZZInit = 0 Then Exit Property

            If mvarDatatype = Displacer Then
                If mvarCyclicRate > 4 Then
                    'mvarCyclicRate = 4
                    'InfoPrint 1, "Minimum cyclic rate for Displacers is 4"
                End If
            Else
                If mvarCyclicRate > 20 Then
                    mvarCyclicRate = 20
                    modHelper.InfoPrint(1, "Maximum cyclic rate is 20")
                ElseIf mvarCyclicRate < 1 / 2 Then  'note: if i set min value to 1/5 make sure i update the wdList
                    'mvarCyclicRate = 1 / 2
                    'InfoPrint 1, "Minimum cyclic rate is 1/2"
                End If
            End If

        End Set
    End Property





    Public Property BeamOutput() As Single
        Get


            BeamOutput = mvarBeamOutput
        End Get
        Set(ByVal Value As Single)



            mvarBeamOutput = Value
            If mvarZZInit = 0 Then Exit Property

            Select Case mvarDatatype

                Case Laser, BlueGreenLaser, RainbowLaser, UVLaser, IRLaser, Disruptor, Flamer, Screamer, Stunner, ParalysisBeam, XRayLaser, FusionBeam, GravityBeam, Graser, Disintegrator, BeamedPowerTransmitter, MilitaryParalysisBeam
                    If mvarBeamOutput < 60 Then
                        mvarBeamOutput = 60
                        modHelper.InfoPrint(1, "Minimum beam output for this weapon is 60 kJ")
                    End If

                Case ChargedParticleBeam, NeutralParticleBeam
                    If mvarTL <= 8 Then
                        If mvarBeamOutput < 10000 Then
                            mvarBeamOutput = 10000
                            modHelper.InfoPrint(1, "Minimum beam output for particle Beams at this TL is 10,000 kJ")
                        End If
                    Else
                        If mvarBeamOutput < 60 Then
                            mvarBeamOutput = 60
                            modHelper.InfoPrint(1, "Minimum beam output for Antiparticle Beams at this TL is 60 kJ")
                        End If
                    End If

                Case AntiparticleBeam
                    If mvarTL <= 12 Then
                        If mvarBeamOutput < 10000 Then
                            mvarBeamOutput = 10000
                            modHelper.InfoPrint(1, "Minimum beam output for Antiparticle Beams at this TL is 10,000 kJ")
                        End If
                    Else
                        If mvarBeamOutput < 60 Then
                            mvarBeamOutput = 60
                            modHelper.InfoPrint(1, "Minimum beam output for Antiparticle Beams at this TL is 60 kJ")
                        End If
                    End If

                Case Displacer
                    If mvarBeamOutput < 24000 Then
                        mvarBeamOutput = 24000
                        modHelper.InfoPrint(1, "Minimum beam output for Displacers is 24,000 kJ")
                    End If

            End Select

        End Set
    End Property
	
	
	
	
	
	Public Property Quality() As String
		Get


            Quality = mvarQuality
        End Get
        Set(ByVal Value As String)


            mvarQuality = Value
        End Set
    End Property



    Public Property CustomDescription() As String
        Get


            CustomDescription = mvarCustomDescription
        End Get
        Set(ByVal Value As String)


            mvarCustomDescription = Value
        End Set
    End Property
	
	
	Public Property Weight() As Double
		Get


            Weight = mvarWeight
        End Get
        Set(ByVal Value As Double)


            mvarWeight = Value
        End Set
    End Property


    Public Property Volume() As Double
        Get


            Volume = mvarVolume
        End Get
        Set(ByVal Value As Double)


            mvarVolume = Value
        End Set
    End Property
	
	
	Public Property TL() As Short
		Get


            TL = mvarTL
        End Get
        Set(ByVal Value As Short)


            mvarTL = Value
        End Set
    End Property



    Public Property SurfaceArea() As Double
        Get


            SurfaceArea = mvarSurfaceArea
        End Get
        Set(ByVal Value As Double)


            mvarSurfaceArea = Value
        End Set
    End Property
	
	
	Public Property Quantity() As Short
		Get


            Quantity = mvarQuantity
        End Get
        Set(ByVal Value As Short)


            mvarQuantity = Value
        End Set
    End Property


    Public Property Parent() As String
        Get


            Parent = mvarParent
        End Get
        Set(ByVal Value As String)


            mvarParent = Value
        End Set
    End Property
	
	
	Public Property Key() As String
		Get


            Key = mvarKey
        End Get
        Set(ByVal Value As String)


            mvarKey = Value
        End Set
    End Property



    Public Property SelectedImage() As Short
        Get


            SelectedImage = mvarSelectedImage
        End Get
        Set(ByVal Value As Short)


            mvarSelectedImage = Value
        End Set
    End Property
	
	
	
	
	
	Public Property Image() As Short
		Get


            Image = mvarImage
        End Get
        Set(ByVal Value As Short)


            mvarImage = Value
        End Set
    End Property





    Public Property Description() As String
        Get


            Description = mvarDescription
        End Get
        Set(ByVal Value As String)


            mvarDescription = Value
        End Set
    End Property
	
	
	
	
	
	Public Property Datatype() As Short
		Get


            Datatype = mvarDatatype
        End Get
        Set(ByVal Value As Short)


            mvarDatatype = Value
        End Set
    End Property





    Public Property Cost() As Double
        Get


            Cost = mvarCost
        End Get
        Set(ByVal Value As Double)


            mvarCost = Value
        End Set
    End Property
	
	
	
	
	
	Public Property Custom() As Boolean
		Get


            Custom = mvarCustom
        End Get
        Set(ByVal Value As Boolean)


            mvarCustom = Value
        End Set
    End Property





    Public Property DR() As Integer
        Get


            DR = mvarDR
        End Get
        Set(ByVal Value As Integer)


            mvarDR = Value
        End Set
    End Property
	
	
	
	
	
	Public Property HitPoints() As Double
		Get


            HitPoints = mvarHitPoints
        End Get
        Set(ByVal Value As Double)


            mvarHitPoints = Value
        End Set
    End Property




    Public Property Ruggedized() As Boolean
        Get


            Ruggedized = mvarRuggedized
        End Get
        Set(ByVal Value As Boolean)


            mvarRuggedized = Value
        End Set
    End Property
	
	
	
	
	Public Function LocationCheck() As Boolean
		Dim TempCheck As Boolean
		Dim InstallPoint As Short
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		InstallPoint = Veh.Components(mvarParent).Datatype
		
		
		If (InstallPoint = Body) Or (InstallPoint = GroupComponent) Or (InstallPoint = Superstructure) Or (InstallPoint = Pod) Or (InstallPoint = Turret) Or (InstallPoint = Popturret) Or (InstallPoint = Arm) Or (InstallPoint = Wing) Or (InstallPoint = OpenMount) Or (InstallPoint = Leg) Or (InstallPoint = equipmentPod) Or (InstallPoint = Module_Renamed) Then
			TempCheck = True
		Else
			modHelper.InfoPrint(1, "Weapons must be placed in Body, Superstructure, Pod, equipment Pod,Turret, Popturret, Arm, Wing, Open Mount, Leg or Module.")
			TempCheck = False
		End If
		
		If TempCheck Then SetLogicalParent()
		LocationCheck = TempCheck
	End Function
	
	
	Private Function GetLocation() As String
		On Error Resume Next
		If mvarLogicalParent = "" Then SetLogicalParent()
		'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetLocation = Veh.Components(mvarLogicalParent).Abbrev
		
	End Function
	
	Public Sub SetLogicalParent()
		mvarLogicalParent = GetLogicalParent(mvarParent)
	End Sub
	
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		
		' set the default properties
		mvarCustom = False
		TL = gVehicleTL
		mvarQuantity = 1
		mvarQuality = "normal"
		mvarMount = "normal"
		mvarBeamOutput = 1600
		mvarCyclicRate = 5
		mvarRange = "normal"
		mvarFTL = False
		mvarCompact = False
		mvarReliable = False
		mvarRoF = CStr(5)
		mvarPowerCellType = "none"
		mvarPowerCellQuantity = 0
		mvarDirection = "front"
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'the class is being destroyed
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
	
	Public Sub Init()
		
		'NOTE: i removed energydrill and beamedpowertransmitter from Database.
		'they'll be added in version 1.1
		
		Select Case mvarDatatype
			Case Laser
				
			Case UVLaser
				
			Case BlueGreenLaser
				
			Case RainbowLaser
				
			Case IRLaser
				
			Case Disruptor
				
			Case ChargedParticleBeam
				
			Case NeutralParticleBeam
				
			Case Flamer
				
			Case Screamer
				
			Case Stunner
				
			Case ParalysisBeam
				
			Case XRayLaser
				
			Case FusionBeam
				
			Case GravityBeam
				
			Case AntiparticleBeam
				
			Case Graser
				
			Case Disintegrator
				
			Case Displacer
				
			Case BeamedPowerTransmitter
				
			Case MilitaryParalysisBeam
				
				
				
		End Select
		
	End Sub
	
	
	Public Sub StatsUpdate()
		Dim sPrint1 As String
		Dim sPrint2 As String
		Dim sPrintPlural As String
		Dim sPrintPlural2 As String
		Dim sPrintPlural3 As String
		Dim sPrintDirection As String
		Dim QRugMod As Single
		Dim RugHitMod As Integer
		
		mvarZZInit = 1
		'calculate statistics
		mvarLocation = GetLocation
		
		'set the ruggedized and quantity multipliers
		If mvarQuantity < 1 Then mvarQuantity = 1
		If mvarRuggedized Then
			QRugMod = 1.5 * mvarQuantity
			RugHitMod = 2
		Else
			QRugMod = 1 * mvarQuantity
			RugHitMod = 1
		End If
		
		
		mvarMalfunction = GetMalfunction
		GetTypeDamages() 'call sub to update Damages based on Ammunition Type
		
		mvarKEDamage = GetDamage
		
		mvarhalfDamage = GetHalfDamage
		If mvarDatatype = Stunner Then
			mvarMaxRange = mvarhalfDamage * 3
			mvarMaxRange2 = mvarhalfDamage * 2
		Else
			mvarMaxRange = GetMaxRange
		End If
		
		'set the ranges in a vacuum
		SetVacuumRanges()
		
		'mvarMinRange = GetMinRange
		mvarAccuracy = GetAccuracy
		mvarWeight = GetWeight
		
		mvarSnapShot = GetSnapShot
		mvarCost = GetCost
		mvarPowerReqt = GetPowerReqt
		mvarShots = GetShots
		mvarWeight = mvarWeight + mvarPowerCellWeight 'update weight to include power cells
		mvarVolume = GetVolume
		mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
		'UPGRADE_WARNING: Couldn't resolve default property of object CalcComponentHitpoints(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		mvarHitPoints = CalcComponentHitpoints(mvarSurfaceArea) * RugHitMod
		'cost, malf, and accuracy modifiers for Cheap, Fine and Very Fine quality are calced in the functions below
		'todo check that quantity is being worked out correctly
		'todo check that "mounting" option is taken into account
		
		
		'//update the cost,weight,volume, surface area and volume based on quantity and ruggedized options
		mvarCost = mvarCost * QRugMod
		mvarWeight = mvarWeight * QRugMod
		mvarVolume = mvarVolume * QRugMod
		
		mvarSurfaceArea = CalcSurfaceArea(mvarVolume)
		mvarPowerReqt = mvarPowerReqt * mvarQuantity
		
		'produce the print output
		'produce the print output
		If mvarRuggedized Then
			sPrint1 = "ruggedized "
		Else
			sPrint1 = ""
		End If
		
		If mvarPowerCellQuantity > 0 Then
			sPrint2 = ", " & mvarPowerCellQuantity & " " & mvarPowerCellType
		End If
		If mvarMount <> "normal" Then
			sPrint2 = sPrint2 & ", " & mvarMount
		End If
		If mvarFTL Then
			sPrint2 = sPrint2 & ", FTL"
		End If
		If mvarCompact Then
			sPrint2 = sPrint2 & ", compact"
		End If
		If mvarReliable Then
			sPrint2 = sPrint2 & ", reliable reputation"
		End If
		If mvarQuality <> "normal" Then sPrint2 = sPrint2 & ", " & mvarQuality & " construction"
		
		
		sPrintDirection = StrConv(Left(mvarDirection, 1), VbStrConv.UpperCase)
		
		If mvarQuantity > 1 Then
			sPrintPlural = "s"
			sPrintPlural2 = " each"
			sPrintPlural3 = " total of "
		Else
			sPrintPlural = ""
			sPrintPlural2 = ""
			sPrintPlural3 = ""
		End If
		
		mvarPrintOutput = NumericToString(mvarQuantity) & " TL" & mvarTL & " " & sPrint1 & mvarCustomDescription & sPrintPlural & sPrint2 & " (" & mvarLocation & sPrintDirection & ", HP " & mvarHitPoints & sPrintPlural2 & ", " & sPrintPlural3 & VB6.Format(mvarWeight, p_sFormat) & " lbs., " & VB6.Format(mvarVolume, p_sFormat) & " cf., " & "$" & VB6.Format(mvarCost, p_sFormat) & ", " & VB6.Format(mvarPowerReqt, p_sFormat) & " kW)." & mvarComment
		
	End Sub
	
	Public Sub QueryParent()
		' if the object has a parent, query it and check to see if
		' more stats/property updates are needed for other objects in the collection
		If mvarParent <> "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object Veh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Veh.Components(Parent).StatsUpdate()
		End If
	End Sub
	
	
	
	Private Function GetMalfunction() As String
		Dim TempMalf As String
		
		Select Case mvarDatatype
			
			Case Disruptor, NeutralParticleBeam, ChargedParticleBeam
				If mvarTL <= 9 Then
					TempMalf = "Crit."
				Else
					TempMalf = "Ver."
				End If
			Case Disintegrator, Displacer
				TempMalf = "Crit."
			Case UVLaser, BlueGreenLaser, IRLaser
				If mvarTL >= 9 Then
					TempMalf = "Ver.(Crit.)"
				Else
					TempMalf = "Ver."
				End If
			Case RainbowLaser
				If mvarTL >= 10 Then
					TempMalf = "Ver.(Crit.)"
				Else
					TempMalf = "Ver."
				End If
			Case Else
				TempMalf = "Ver."
				
		End Select
		
		'increase malf if weapon has reputation for quality
		If mvarReliable Then IncreaseMalf(TempMalf)
		
		'get modifier for Cheap, Fine and Very Fine quality
		If mvarQuality = "cheap" Then
			TempMalf = DecreaseMalf(TempMalf)
		ElseIf mvarQuality = "fine (reliable)" Then 
			TempMalf = IncreaseMalf(TempMalf)
		End If
		
		GetMalfunction = TempMalf
	End Function
	
	
	Sub GetTypeDamages()
		
		Select Case mvarDatatype
			
			Case Laser, BlueGreenLaser, RainbowLaser, UVLaser, IRLaser, ChargedParticleBeam, NeutralParticleBeam, XRayLaser, GravityBeam, Graser, BeamedPowerTransmitter
				
				mvarTypeDamage = "Imp."
				
			Case AntiparticleBeam, Disruptor, Flamer, Screamer, Stunner, FusionBeam, ParalysisBeam, MilitaryParalysisBeam, Disintegrator, Displacer
				mvarTypeDamage = "Spcl."
		End Select
		
	End Sub
	Private Function GetDamage() As Single
		Dim fKEDamage As Single 'holds numeric value of Kinetic Energy damage before its converted to a GURPS format string
		Dim Suffix As String 'suffix for armor divisor
		Dim O As Single 'beam output
		Dim b As Single 'beam type  modifier
		Dim E As Single 'beam energy output modifier
		Dim T As Single 'beam tech level modifier
		Dim Diff As Single
		Dim TempDamage As String
		
		'get suffix
		If (mvarDatatype = XRayLaser) Or (mvarDatatype = AntiparticleBeam) Then
			Suffix = "(2)"
		ElseIf mvarDatatype = Graser Then 
			Suffix = "(5)"
		ElseIf (mvarDatatype = Disintegrator) Or (mvarDatatype = GravityBeam) Then 
			Suffix = "(100)"
		Else
			Suffix = ""
		End If
		
		
		'Get O
		O = mvarBeamOutput
		
		'E not used in Errata
		'get E
		'If (mvarDatatype = Screamer) Or (mvarDatatype = Stunner) Or (mvarDatatype = MilitaryParalysisBeam) Or (mvarDatatype = ParalysisBeam) Then
		'    E = 1
		'Else
		'    If mvarBeamOutput < 1000 Then
		'        E = 0.5
		'    ElseIf mvarBeamOutput < 200 Then
		'        E = 0.33
		'    Else
		'        E = 1
		'    End If
		'End If
		
		'get T and B
		Select Case mvarDatatype
			Case Laser
				b = 0.5
				Diff = 8
			Case UVLaser
				b = 0.5
				Diff = 8
			Case BlueGreenLaser
				b = 0.5
				Diff = 8
			Case RainbowLaser
				b = 0.5
				Diff = 9
			Case IRLaser
				b = 0.5
				Diff = 8
			Case Disruptor
				b = 1.6
				Diff = 8
			Case ChargedParticleBeam
				b = 1.6
				Diff = 8
			Case NeutralParticleBeam
				b = 1.6
				Diff = 8
			Case Flamer
				b = 2
				Diff = 9
			Case Screamer
				b = 2.5
				Diff = 9
			Case Stunner
				b = 0.4 'changed to match errata
				Diff = 9
			Case ParalysisBeam
				b = 0.3 'changed to match errata
				Diff = 10
			Case XRayLaser
				b = 0.5
				Diff = 10
			Case FusionBeam
				b = 4
				Diff = 12
			Case GravityBeam
				b = 0.04
				Diff = 12
			Case AntiparticleBeam
				b = 2
				Diff = 12
			Case Graser
				b = 0.5
				Diff = 14
			Case Disintegrator
				b = 0.36
				Diff = 15
			Case Displacer
				b = 0.066
				Diff = 16
			Case BeamedPowerTransmitter
				Diff = 8
			Case MilitaryParalysisBeam
				b = 0.3 'changed to match errata
				Diff = 10
				
		End Select
		
		Diff = mvarTL - Diff
		
		If Diff <= 0 Then
			T = 1
		ElseIf Diff = 1 Then 
			T = 1.2857
		ElseIf Diff = 2 Then 
			T = 1.5714
		ElseIf Diff >= 3 Then 
			T = 1.8571
		End If
		
		'dice of damage formula
		If O < 1000 Then
			fKEDamage = (O * 0.03) * b * T
		ElseIf O > 1000000 Then 
			fKEDamage = (10 * (O ^ (1 / 3))) * b * T
		Else
			fKEDamage = System.Math.Sqrt(O) * b * T
		End If
		
		'since not all weapons use straight formula,perform the exceptions
		If (mvarDatatype = Stunner) Or (mvarDatatype = MilitaryParalysisBeam) Then
			mvarDamage = CStr(-(System.Math.Round(fKEDamage, 0)))
		ElseIf mvarDatatype = ParalysisBeam Then 
			mvarDamage = CStr(-(System.Math.Round(fKEDamage, 0)) + 4)
		ElseIf mvarDatatype = Displacer Then 
			mvarDamage = CStr(fKEDamage) 'this is blast radius in yards
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object ConvertDamage(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			TempDamage = ConvertDamage(fKEDamage)
			If TempDamage = "No Damage" Then
				mvarDamage = TempDamage
			Else
				mvarDamage = TempDamage & Suffix
			End If
		End If
		
		GetDamage = fKEDamage
		
	End Function
	
	Private Function GetHalfDamage() As Single 'in yards
		Dim TempHalfDamage As Double
		Dim R As Single 'range modifier
		Dim b As Short 'datatype modifier
		Dim T As Single 'tech level modifier
		Dim Diff As Short
		
		'some beam types have no 1/2 damage range
		If (mvarDatatype = ParalysisBeam) Or (mvarDatatype = Disintegrator) Or (mvarDatatype = Displacer) Or (mvarDatatype = MilitaryParalysisBeam) Then
			GetHalfDamage = 0
			Exit Function
		End If
		
		'get R
		If mvarRange = "extreme" Then
			R = 8
		ElseIf mvarRange = "very long" Then 
			R = 4
		ElseIf mvarRange = "long" Then 
			R = 2
		ElseIf mvarRange = "close" Then 
			R = 0.25
		ElseIf mvarRange = "normal" Then 
			R = 1
		End If
		
		'get B
		If (mvarDatatype = Stunner) Or (mvarDatatype = FusionBeam) Then
			b = 12
		ElseIf (mvarDatatype = Flamer) Or (mvarDatatype = Screamer) Then 
			b = 18
		ElseIf mvarDatatype = GravityBeam Then 
			b = 36
		ElseIf (mvarDatatype = ChargedParticleBeam) Or (mvarDatatype = NeutralParticleBeam) Or (mvarDatatype = AntiparticleBeam) Or (mvarDatatype = Disruptor) Then 
			b = 15
		ElseIf (mvarDatatype = BlueGreenLaser) Or (mvarDatatype = RainbowLaser) Or (mvarDatatype = Laser) Then 
			b = 200
		ElseIf mvarDatatype = IRLaser Then 
			b = 170 'added for errata
		ElseIf mvarDatatype = UVLaser Then 
			b = 60
		ElseIf (mvarDatatype = XRayLaser) Or (mvarDatatype = Graser) Then 
			b = 75
		End If
		
		'get T
		Select Case mvarDatatype
			Case Laser
				Diff = 8
			Case UVLaser
				Diff = 8
			Case BlueGreenLaser
				Diff = 8
			Case RainbowLaser
				Diff = 9
			Case IRLaser
				Diff = 8
			Case Disruptor
				Diff = 8
			Case ChargedParticleBeam
				Diff = 8
			Case NeutralParticleBeam
				Diff = 8
			Case Flamer
				Diff = 9
			Case Screamer
				Diff = 9
			Case Stunner
				Diff = 9
			Case XRayLaser
				Diff = 10
			Case FusionBeam
				Diff = 12
			Case GravityBeam
				Diff = 12
			Case AntiparticleBeam
				Diff = 12
			Case Graser
				Diff = 14
			Case BeamedPowerTransmitter
				Diff = 8
				
		End Select
		
		Diff = mvarTL - Diff
		
		If Diff <= 0 Then
			T = 1
		ElseIf Diff = 1 Then 
			T = 1.1
		ElseIf Diff = 2 Then 
			T = 1.2
		ElseIf Diff >= 3 Then 
			T = 1.3
		End If
		
		'calc formula
		TempHalfDamage = System.Math.Sqrt(mvarBeamOutput) * R * b * T
		
		GetHalfDamage = RoundRange(TempHalfDamage)
	End Function
	
	Private Function GetMaxRange() As Double 'in yards
		
		Dim TempMax As Double
		
		Dim O As Single
		Dim R As Single
		Dim b As Single
		Dim T As Single
		Dim Diff As Short
		Select Case mvarDatatype
			
			Case Laser, BlueGreenLaser, RainbowLaser, UVLaser, IRLaser, Disruptor, Screamer, XRayLaser, Graser, BeamedPowerTransmitter
				
				If mvarKEDamage >= 20 Then
					TempMax = mvarhalfDamage * 3
				Else
					TempMax = mvarhalfDamage * 2
				End If
				
			Case AntiparticleBeam, ChargedParticleBeam, NeutralParticleBeam
				If mvarKEDamage >= 100 Then
					TempMax = mvarhalfDamage * 3
				Else
					TempMax = mvarhalfDamage * 2
				End If
				
			Case Flamer, FusionBeam
				If mvarKEDamage >= 10 Then
					TempMax = mvarhalfDamage * 3
				Else
					TempMax = mvarhalfDamage * 2
				End If
				
				
			Case Stunner
				'this is taken care of in the statsupdate routine in this class module
				
				
			Case GravityBeam
				If mvarKEDamage >= 3 Then
					TempMax = mvarhalfDamage * 3
				Else
					TempMax = mvarhalfDamage * 2
				End If
				
			Case ParalysisBeam, MilitaryParalysisBeam, Disintegrator, Displacer
				
				'get O
				O = mvarBeamOutput
				
				'get R
				If mvarRange = "extreme" Then
					R = 8
				ElseIf mvarRange = "very long" Then 
					R = 4
				ElseIf mvarRange = "long" Then 
					R = 2
				ElseIf mvarRange = "close" Then 
					R = 0.25
				ElseIf mvarRange = "normal" Then 
					R = 1
				End If
				
				'get B
				If (mvarDatatype = Displacer) Then
					b = 1.3
				ElseIf (mvarDatatype = ParalysisBeam) Or (mvarDatatype = MilitaryParalysisBeam) Then 
					b = 24
				ElseIf (mvarDatatype = Disintegrator) Then 
					b = 70
				End If
				
				'get T
				Select Case mvarDatatype
					Case ParalysisBeam, MilitaryParalysisBeam
						Diff = 10
					Case Disintegrator
						Diff = 15
					Case Displacer
						Diff = 16
				End Select
				
				Diff = mvarTL - Diff
				
				If Diff <= 0 Then
					T = 1
				ElseIf Diff = 1 Then 
					T = 1.1
				ElseIf Diff = 2 Then 
					T = 1.2
				ElseIf Diff >= 3 Then 
					T = 1.3
				End If
				
				'calc formula
				TempMax = System.Math.Sqrt(mvarBeamOutput) * R * b * T
		End Select
		
		
		GetMaxRange = RoundRange(TempMax)
	End Function
	
	Private Function SetVacuumRanges() As Single 'in yards
		Dim Modifier As Single
		
		Select Case mvarDatatype
			
			Case UVLaser, RainbowLaser, NeutralParticleBeam
				Modifier = 50
				
			Case Laser, BlueGreenLaser, IRLaser, Flamer, Disruptor, AntiparticleBeam, BeamedPowerTransmitter
				Modifier = 10
				
			Case ChargedParticleBeam
				Modifier = 0.01
				
			Case XRayLaser, Graser
				Modifier = 100
				
			Case Stunner, Screamer
				'this is taken care of in the statsupdate routine in this class module
				Modifier = 0
				
			Case GravityBeam, FusionBeam, ParalysisBeam, MilitaryParalysisBeam, Disintegrator, Displacer
				Modifier = 1
		End Select
		
		mvarVacuumHalfDamage = Modifier * mvarhalfDamage
		mvarVacuumMaxRange = Modifier * mvarMaxRange
		
	End Function
	
	Private Function RoundRange(ByRef Range As Double) As Double
		
		' Do the final rounding
		If Range <= 100 Then
			' round to nearest yard
			Range = System.Math.Round(Range, 0)
		ElseIf Range <= 1000 Then  ' round to nearest 10 yards
			Range = System.Math.Round(Range / 10, 0) * 10
		ElseIf Range <= 10000 Then  ' round to nearest 100 yards
			Range = System.Math.Round(Range / 100, 0) * 100
		ElseIf Range > 10000 Then  ' round to nearest 1000 yards
			Range = System.Math.Round(Range / 1000, 0) * 1000
		End If
		
		RoundRange = Range
	End Function
	
	Private Function GetAccuracy() As Integer
		
		
		Dim R As Single
		Dim Acc As Short
		Dim Base As Short
		Dim i As Integer
		
		'get R
		If (mvarDatatype = ParalysisBeam) Or (mvarDatatype = MilitaryParalysisBeam) Or (mvarDatatype = Disintegrator) Or (mvarDatatype = Displacer) Then
			R = mvarMaxRange
		Else
			R = mvarhalfDamage
		End If
		
		'find acc
		Acc = 0
		Base = 11
		i = 0
		
		Do While Acc = 0
			If R <= 150 Then 'this is the only exception to this formula
				Acc = 11
				Exit Do
			End If
			If R < 150 * 10 ^ i Then
				Acc = Base + (6 * i)
			ElseIf R < 200 * 10 ^ i Then 
				Acc = Base + 1 + (6 * i)
			ElseIf R < 300 * 10 ^ i Then 
				Acc = Base + 2 + (6 * i)
			ElseIf R < 450 * 10 ^ i Then 
				Acc = Base + 3 + (6 * i)
			ElseIf R < 700 * 10 ^ i Then 
				Acc = Base + 4 + (6 * i)
			ElseIf R < 1000 * 10 ^ i Then 
				Acc = Base + 5 + (6 * i)
			End If
			i = i + 1
		Loop 
		
		'//if energy drill, reduce acc by 5
		If mvarEnergyDrill Then
			Acc = Acc - 5
		End If
		
		'get modifier for Cheap, Fine and Very Fine quality
		If mvarQuality = "cheap" Then
			Acc = Acc - 1
		ElseIf mvarQuality = "fine (accurate)" Then 
			Acc = Acc + 1
		ElseIf mvarQuality = "very fine (accurate)" Then 
			Acc = Acc + 2
		End If
		
		'note: should this be done before or after cheap,fine calcs above?
		'modifiers for certain types of weapons
		If (mvarDatatype = ParalysisBeam) Or (mvarDatatype = MilitaryParalysisBeam) Or (mvarDatatype = Disintegrator) Or (mvarDatatype = Displacer) Then
			Acc = Acc - 2
		ElseIf mvarDatatype = Flamer Then 
			If Acc < 20 Then Acc = 20
		End If
		
		GetAccuracy = Acc
	End Function
	
	Private Function GetWeight() As Double
		Dim O As Single
		Dim b As Integer
		Dim s As Single
		Dim l As Single
		Dim F As Single
		Dim R As Single
		Dim TempWeight As Double
		
		'Get o
		O = mvarBeamOutput
		
		'get B
		Select Case mvarDatatype
			
			Case Stunner, Flamer
				b = 12
			Case BeamedPowerTransmitter, Screamer, Laser, BlueGreenLaser, UVLaser, IRLaser
				b = 24
			Case RainbowLaser
				b = 32
			Case XRayLaser
				b = 72
			Case NeutralParticleBeam
				b = 64
			Case Graser
				b = 128
			Case ChargedParticleBeam
				b = 72
			Case Disruptor
				b = 36
			Case ParalysisBeam, MilitaryParalysisBeam
				b = 80
			Case AntiparticleBeam, FusionBeam
				b = 240
			Case Disintegrator, GravityBeam
				b = 180
			Case Displacer
				b = 1350
		End Select
		
		'get S
		If O <= 6400 Then
			s = 0.5
		Else
			s = 1
		End If
		
		'get L
		If mvarCompact Then
			l = 0.5
		Else
			l = 1
		End If
		
		'get F
		If mvarCyclicRate <= 1 / 2 Then
			F = 0.666
		ElseIf mvarCyclicRate = 1 Then 
			F = 1
		ElseIf mvarCyclicRate = 2 Then 
			F = 1.25
		ElseIf mvarCyclicRate = 3 Then 
			F = 1.5
		ElseIf mvarCyclicRate >= 4 Then 
			Select Case mvarDatatype 'todo: should this exlclude any laser types?
				Case Laser, BlueGreenLaser, RainbowLaser, UVLaser, IRLaser, XRayLaser, Graser, BeamedPowerTransmitter
					
					F = mvarCyclicRate / 2
				Case Else
					F = mvarCyclicRate / 4 + 1
			End Select
		End If
		
		'get R
		If mvarRange = "extreme" Then
			R = 4
		ElseIf mvarRange = "very long" Then 
			R = 2
		ElseIf mvarRange = "long" Then 
			R = 1.5
		ElseIf mvarRange = "close" Then 
			R = 0.666
		ElseIf mvarRange = "normal" Then 
			R = 1
		End If
		
		
		TempWeight = (O / b) * s * l * F * R
		
		
		GetWeight = System.Math.Round(TempWeight, 2)
		
	End Function
	
	Private Function GetVolume() As Double
		If mvarMount = "normal" Then
			GetVolume = mvarWeight / 50
		Else
			GetVolume = mvarWeight / 20 'concealed weapons take up more space
		End If
		
	End Function
	
	Private Function GetSnapShot() As Integer
		Dim TSS As Integer
		
		If mvarWeight < 2.5 Then
			TSS = 11
		ElseIf mvarWeight < 10 Then 
			TSS = 12
		ElseIf mvarWeight < 15 Then 
			TSS = 14
		ElseIf mvarWeight < 26 Then 
			TSS = 17
		ElseIf mvarWeight < 401 Then 
			TSS = 20
		ElseIf mvarWeight < 2001 Then 
			TSS = 25
		Else
			TSS = 30
		End If
		
		'half SS for flamers (pg 126 bottom)
		If mvarDatatype = Flamer Then TSS = TSS / 2
		
		'energy drills are SS + 5
		If mvarEnergyDrill Then
			TSS = TSS + 5
		End If
		
		GetSnapShot = TSS
	End Function
	
	
	
	Private Function GetCost() As Double
		Dim TempCost As Double
		Dim CompactMod As Short
		Dim b As Single
		Dim Diff As Short
		Dim T As Short
		
		'get basic Cost
		If mvarWeight < 10 Then
			TempCost = mvarWeight * 200 + 1000
		ElseIf mvarWeight <= 100 Then 
			TempCost = mvarWeight * 300
		ElseIf mvarWeight > 100 Then 
			TempCost = mvarWeight * 100 + 20000
		End If
		
		'get B
		Select Case mvarDatatype
			Case BeamedPowerTransmitter, Laser, BlueGreenLaser, UVLaser, IRLaser, ParalysisBeam
				b = 0.5
			Case RainbowLaser, XRayLaser
				b = 0.666
			Case GravityBeam, Graser
				b = 1.333
			Case Displacer
				b = 4
			Case AntiparticleBeam
				b = 7.5
			Case FusionBeam
				b = 6
			Case Disintegrator
				b = 3.5
			Case MilitaryParalysisBeam, Disruptor, Flamer, Screamer, Stunner, ChargedParticleBeam, NeutralParticleBeam
				b = 1
		End Select
		
		'get CompactMod
		If mvarCompact Then
			CompactMod = 4
		Else
			CompactMod = 1
		End If
		
		'get TL modifier
		'get T
		Select Case mvarDatatype
			Case Laser
				Diff = 8
			Case UVLaser
				Diff = 8
			Case BlueGreenLaser
				Diff = 8
			Case RainbowLaser
				Diff = 9
			Case IRLaser
				Diff = 8
			Case Disruptor
				Diff = 8
			Case ChargedParticleBeam
				Diff = 8
			Case NeutralParticleBeam
				Diff = 8
			Case Flamer
				Diff = 9
			Case Screamer
				Diff = 9
			Case Stunner
				Diff = 9
			Case XRayLaser
				Diff = 10
			Case FusionBeam
				Diff = 12
			Case GravityBeam
				Diff = 12
			Case AntiparticleBeam
				Diff = 12
			Case Graser
				Diff = 14
			Case BeamedPowerTransmitter
				Diff = 8
				
		End Select
		
		Diff = mvarTL - Diff
		
		
		'apply modifiers for Beam Type and Compact option
		TempCost = TempCost * b * CompactMod
		
		'apply modifier for beam TL first appearance
		If Diff <= 0 Then
			T = 1
		ElseIf Diff = 1 Then 
			T = 2
		ElseIf Diff >= 2 Then 
			T = 4
		End If
		
		TempCost = TempCost / T
		
		'//energy drills cost half
		If mvarEnergyDrill Then
			TempCost = TempCost / 2
		End If
		
		'apply modifier for Cheap, Fine and Very Fine quality
		If mvarQuality = "cheap" Then
			TempCost = TempCost / 2
		ElseIf mvarQuality = "fine (accurate)" Then 
			TempCost = TempCost * 5
		ElseIf mvarQuality = "very fine (accurate)" Then 
			TempCost = TempCost * 30
		ElseIf mvarQuality = "fine (reliable)" Then 
			TempCost = TempCost * 5
		End If
		
		GetCost = System.Math.Round(TempCost, 2)
	End Function
	
	
	Private Function GetPowerReqt() As Double
		Dim TempPower As Double
		Dim O As Single
		Dim E As Single
		
		O = mvarBeamOutput
		
		'get E
		Select Case mvarDatatype
			Case GravityBeam, Disruptor, Laser, BlueGreenLaser, UVLaser, IRLaser, BeamedPowerTransmitter, ChargedParticleBeam, NeutralParticleBeam
				E = 2
			Case RainbowLaser
				E = 2.25
			Case AntiparticleBeam, Screamer, Stunner, Disintegrator
				E = 2.5
			Case XRayLaser
				E = 2.666
			Case Graser, ParalysisBeam, MilitaryParalysisBeam
				E = 3
			Case Flamer
				E = 5
			Case FusionBeam
				E = 9
			Case Displacer
				E = 10
		End Select
		
		TempPower = O * E * mvarCyclicRate
		
		GetPowerReqt = System.Math.Round(TempPower, 0)
		
	End Function
	
	Private Function GetShots() As String
		'number of shots the weapon has ready to fire.
		Dim EPS As Single 'equal to Pow/RoF or cyclic rate
		Dim P As Integer
		Dim Multiplier As Short
		Dim NumShots As Single
		Dim Cell As String
		Dim WeightMod As Single
		
		'get EPS
		EPS = mvarPowerReqt / mvarCyclicRate
		
		'get P
		If mvarTL <= 8 Then
			P = 3600
		ElseIf mvarTL = 9 Then 
			P = 5400
		ElseIf mvarTL = 10 Then 
			P = 7200
		ElseIf mvarTL = 11 Then 
			P = 9000
		ElseIf mvarTL = 12 Then 
			P = 10800
		ElseIf mvarTL = 13 Then 
			P = 12600
		ElseIf mvarTL = 14 Then 
			P = 14400
		ElseIf mvarTL = 15 Then 
			P = 16200
		ElseIf mvarTL >= 16 Then 
			P = 18000
		End If
		
		'get Divisor
		If (mvarPowerCellType = "C cells") Or (mvarPowerCellType = "rC cell") Then
			Multiplier = 1
			Cell = "C"
			WeightMod = 0.5
		ElseIf (mvarPowerCellType = "D cells") Or (mvarPowerCellType = "rD cells") Then 
			Multiplier = 10
			Cell = "D"
			WeightMod = 5
		ElseIf (mvarPowerCellType = "E cells") Or (mvarPowerCellType = "rE cells") Then 
			Multiplier = 100
			Cell = "E"
			WeightMod = 20
		End If
		
		'todo: check rules?  round NumShots up?
		NumShots = mvarPowerCellQuantity * Multiplier / (EPS / P)
		
		'if rechargeable, they half shots due to half capacity
		If Left(mvarPowerCellType, 1) = "r" Then
			NumShots = NumShots / 2
		End If
		'set the Weight of the Shots
		mvarPowerCellWeight = mvarPowerCellQuantity * WeightMod
		
		
		GetShots = NumShots & "/" & mvarPowerCellQuantity & Cell
	End Function
End Class