Option Strict Off
Option Explicit On
Module modTablePropulsion
	
	Public Structure udtHarness
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim FlyingMultiplier As Single
		Dim SwimmingMultiplier As Single
		Dim Cost As Single
		Dim Volume As Single
		Dim Efficiency As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtSail
		Dim ID As Short
		Dim TL As Short
		Dim WeightCloth As Single
		Dim WeightSynthetic As Single
		Dim WeightBioplas As Single
		Dim MotiveThrust As Single
		Dim CostCloth As Single
		Dim CostSynthetic As Single
		Dim CostBioplas As Single
		Dim Volume As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtRowing
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim RowingMultiplier As Single
		Dim Cost As Single
		Dim Volume As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtlightSail
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim ThrustSquareModifier As Single
		Dim Cost As Single
		Dim Volume As Single
		Dim Thrust As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtDrivetrain
		Dim ID As Short
		Dim TL As Short
		Dim Weight1 As Single
		Dim Weight2 As Single
		Dim Weight3 As Single
		Dim Cost As Single
		Dim Volume As Single
	End Structure
	
	Public Structure udtAquatic
		Dim ID As Short
		Dim TL As Short
		Dim Weight1 As Single
		Dim Weight2 As Single
		Dim Weight3 As Single
		Dim MotiveThrust As Single
		Dim Cost As Single
		Dim Volume As Single
	End Structure
	
	Public Structure udtAirscrew
		Dim ID As Short
		Dim TL As Short
		Dim Weight1 As Single
		Dim Weight2 As Single
		Dim Weight3 As Single
		Dim MotiveThrust As Single
		Dim Cost As Single
		Dim Volume As Single
	End Structure
	
	Public Structure udtHelicopter
		Dim ID As Short
		Dim TL As Short
		Dim Lift As Single
		Dim MotiveThrust As Single
		Dim Weight1 As Single
		Dim Weight2 As Single
		Dim Weight3 As Single
		Dim Cost As Single
		Dim Volume As Single
	End Structure
	
	Public Structure udtOrnithopter
		Dim ID As Short
		Dim TL As Short
		Dim MotiveThrust As Single
		Dim Lift As Single
		Dim Weight1 As Single
		Dim Weight2 As Single
		Dim Weight3 As Single
		Dim Cost As Single
		Dim Volume As Single
	End Structure
	
	Public Structure udtJet
		Dim ID As Short
		Dim TL As Short
		Dim Weight1 As Single
		Dim Weight2 As Single
		Dim Cost As Single
		Dim Volume As Single
		Dim ABWeight As Single
		Dim ABCost As Single
		Dim ABThrust As Single
		Dim ABFuel As Single
		Dim Fuel As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtRocket
		Dim ID As Short
		Dim TL As Short
		Dim Weight1 As Single
		Dim Weight2 As Single
		Dim Cost As Single
		Dim Volume As Single
		Dim Fuel As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtOrionEngine
		Dim ID As Short
		Dim TL As Short
		Dim Weight1 As Single
		Dim Weight2 As Single
		Dim Weight3 As Single
		Dim MotiveThrust As Single
		Dim Volume As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtThrustBomb
		Dim ID As Short
		Dim TL As Short
		Dim Weight1 As Single
		Dim Weight2 As Single
		Dim Cost As Single
		Dim Volume As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtSolidRocket
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Cost As Single
		Dim Volume As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtReactionless
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Cost As Single
		Dim MinCost As Single
		Dim Volume As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtMagLev
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Cost As Single
		Dim Volume As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtStarDrve
		Dim ID As Short
		Dim TL As Short
		Dim Weight1 As Single
		Dim Weight2 As Single
		Dim Cost1 As Single
		Dim Cost2 As Single
		Dim Volume As Single
		Dim Power1 As Single
		Dim Power2 As Single
	End Structure
	
	Public Structure udtLiftingGas
		Dim ID As Short
		Dim TL As Short
		Dim Cost As Single
		Dim Volume As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtContraGravity
		Dim ID As Short
		Dim TL As Short
		Dim Weight1 As Single
		Dim Weight2 As Single
		Dim Cost1 As Single
		Dim Cost2 As Single
		Dim Volume As Single
		Dim Power As Single
	End Structure
	
	Public HarnessMatrix() As udtHarness
	Public SailMatrix() As udtSail
	Public RowingMatrix() As udtRowing
	Public lightSailMatrix() As udtlightSail
	Public DrivetrainMatrix() As udtDrivetrain
	Public AquaticMatrix() As udtAquatic
	Public AirscrewMatrix() As udtAirscrew
	Public HelicopterMatrix() As udtHelicopter
	Public OrnithopterMatrix() As udtOrnithopter
	Public JetMatrix() As udtJet
	Public RocketMatrix() As udtRocket
	Public OrionEngineMatrix() As udtOrionEngine
	Public ThrustBombMatrix() As udtThrustBomb
	Public SolidRocketMatrix() As udtSolidRocket
	Public ReactionlessMatrix() As udtReactionless
	Public MagLevMatrix() As udtMagLev
	Public StarDriveMatrix() As udtStarDrve
	Public LiftingGasMatrix() As udtLiftingGas
	Public ContraGravMatrix() As udtContraGravity
	
	Sub LoadHarness()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\1001.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve HarnessMatrix(i)
			Input(10, HarnessMatrix(i).ID)
			Input(10, HarnessMatrix(i).TL)
			Input(10, HarnessMatrix(i).Weight)
			Input(10, HarnessMatrix(i).FlyingMultiplier)
			Input(10, HarnessMatrix(i).SwimmingMultiplier)
			Input(10, HarnessMatrix(i).Cost)
			Input(10, HarnessMatrix(i).Volume)
			Input(10, HarnessMatrix(i).Efficiency)
			Input(10, HarnessMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadSail()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\1002.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve SailMatrix(i)
			Input(10, SailMatrix(i).ID)
			Input(10, SailMatrix(i).TL)
			Input(10, SailMatrix(i).WeightCloth)
			Input(10, SailMatrix(i).WeightSynthetic)
			Input(10, SailMatrix(i).WeightBioplas)
			Input(10, SailMatrix(i).MotiveThrust)
			Input(10, SailMatrix(i).CostCloth)
			Input(10, SailMatrix(i).CostSynthetic)
			Input(10, SailMatrix(i).CostBioplas)
			Input(10, SailMatrix(i).Volume)
			Input(10, SailMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadRowing()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\1003.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve RowingMatrix(i)
			Input(10, RowingMatrix(i).ID)
			Input(10, RowingMatrix(i).TL)
			Input(10, RowingMatrix(i).Weight)
			Input(10, RowingMatrix(i).RowingMultiplier)
			Input(10, RowingMatrix(i).Cost)
			Input(10, RowingMatrix(i).Volume)
			Input(10, RowingMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadlightSail()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\1004.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve lightSailMatrix(i)
			Input(10, lightSailMatrix(i).ID)
			Input(10, lightSailMatrix(i).TL)
			Input(10, lightSailMatrix(i).Weight)
			Input(10, lightSailMatrix(i).ThrustSquareModifier)
			Input(10, lightSailMatrix(i).Cost)
			Input(10, lightSailMatrix(i).Volume)
			Input(10, lightSailMatrix(i).Thrust)
			Input(10, lightSailMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	
	Sub LoadDrivetrain()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\1005.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve DrivetrainMatrix(i)
			Input(10, DrivetrainMatrix(i).ID)
			Input(10, DrivetrainMatrix(i).TL)
			Input(10, DrivetrainMatrix(i).Weight1)
			Input(10, DrivetrainMatrix(i).Weight2)
			Input(10, DrivetrainMatrix(i).Weight3)
			Input(10, DrivetrainMatrix(i).Cost)
			Input(10, DrivetrainMatrix(i).Volume) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadAquatic()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\1006.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve AquaticMatrix(i)
			Input(10, AquaticMatrix(i).ID)
			Input(10, AquaticMatrix(i).TL)
			Input(10, AquaticMatrix(i).Weight1)
			Input(10, AquaticMatrix(i).Weight2)
			Input(10, AquaticMatrix(i).Weight3)
			Input(10, AquaticMatrix(i).MotiveThrust)
			Input(10, AquaticMatrix(i).Cost)
			Input(10, AquaticMatrix(i).Volume) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadAirscrew()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\1007.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve AirscrewMatrix(i)
			Input(10, AirscrewMatrix(i).ID)
			Input(10, AirscrewMatrix(i).TL)
			Input(10, AirscrewMatrix(i).Weight1)
			Input(10, AirscrewMatrix(i).Weight2)
			Input(10, AirscrewMatrix(i).Weight3)
			Input(10, AirscrewMatrix(i).MotiveThrust)
			Input(10, AirscrewMatrix(i).Cost)
			Input(10, AirscrewMatrix(i).Volume) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadHelicopter()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\1008.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve HelicopterMatrix(i)
			Input(10, HelicopterMatrix(i).ID)
			Input(10, HelicopterMatrix(i).TL)
			Input(10, HelicopterMatrix(i).Lift)
			Input(10, HelicopterMatrix(i).MotiveThrust)
			Input(10, HelicopterMatrix(i).Weight1)
			Input(10, HelicopterMatrix(i).Weight2)
			Input(10, HelicopterMatrix(i).Weight3)
			Input(10, HelicopterMatrix(i).Cost)
			Input(10, HelicopterMatrix(i).Volume) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadOrnithopter()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\1009.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve OrnithopterMatrix(i)
			Input(10, OrnithopterMatrix(i).ID)
			Input(10, OrnithopterMatrix(i).TL)
			Input(10, OrnithopterMatrix(i).MotiveThrust)
			Input(10, OrnithopterMatrix(i).Lift)
			Input(10, OrnithopterMatrix(i).Weight1)
			Input(10, OrnithopterMatrix(i).Weight2)
			Input(10, OrnithopterMatrix(i).Weight3)
			Input(10, OrnithopterMatrix(i).Cost)
			Input(10, OrnithopterMatrix(i).Volume) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadJet()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\1010.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve JetMatrix(i)
			Input(10, JetMatrix(i).ID)
			Input(10, JetMatrix(i).TL)
			Input(10, JetMatrix(i).Weight1)
			Input(10, JetMatrix(i).Weight2)
			Input(10, JetMatrix(i).Cost)
			Input(10, JetMatrix(i).Volume)
			Input(10, JetMatrix(i).ABWeight)
			Input(10, JetMatrix(i).ABCost)
			Input(10, JetMatrix(i).ABThrust)
			Input(10, JetMatrix(i).ABFuel)
			Input(10, JetMatrix(i).Fuel)
			Input(10, JetMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadRocket()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\1011.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve RocketMatrix(i)
			Input(10, RocketMatrix(i).ID)
			Input(10, RocketMatrix(i).TL)
			Input(10, RocketMatrix(i).Weight1)
			Input(10, RocketMatrix(i).Weight2)
			Input(10, RocketMatrix(i).Cost)
			Input(10, RocketMatrix(i).Volume)
			Input(10, RocketMatrix(i).Fuel)
			Input(10, RocketMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	
	Sub LoadOrionEngine()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\1012.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve OrionEngineMatrix(i)
			Input(10, OrionEngineMatrix(i).ID)
			Input(10, OrionEngineMatrix(i).TL)
			Input(10, OrionEngineMatrix(i).Weight1)
			Input(10, OrionEngineMatrix(i).Weight2)
			Input(10, OrionEngineMatrix(i).Weight3)
			Input(10, OrionEngineMatrix(i).MotiveThrust)
			Input(10, OrionEngineMatrix(i).Volume)
			Input(10, OrionEngineMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadThrustBomb()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\1013.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve ThrustBombMatrix(i)
			Input(10, ThrustBombMatrix(i).ID)
			Input(10, ThrustBombMatrix(i).TL)
			Input(10, ThrustBombMatrix(i).Weight1)
			Input(10, ThrustBombMatrix(i).Weight2)
			Input(10, ThrustBombMatrix(i).Cost)
			Input(10, ThrustBombMatrix(i).Volume)
			Input(10, ThrustBombMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadSolidRocket()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\1014.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve SolidRocketMatrix(i)
			Input(10, SolidRocketMatrix(i).ID)
			Input(10, SolidRocketMatrix(i).TL)
			Input(10, SolidRocketMatrix(i).Weight)
			Input(10, SolidRocketMatrix(i).Cost)
			Input(10, SolidRocketMatrix(i).Volume)
			Input(10, SolidRocketMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadReactionless()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\1015.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve ReactionlessMatrix(i)
			Input(10, ReactionlessMatrix(i).ID)
			Input(10, ReactionlessMatrix(i).TL)
			Input(10, ReactionlessMatrix(i).Weight)
			Input(10, ReactionlessMatrix(i).Cost)
			Input(10, ReactionlessMatrix(i).MinCost)
			Input(10, ReactionlessMatrix(i).Volume)
			Input(10, ReactionlessMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadMagLev()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\1016.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve MagLevMatrix(i)
			Input(10, MagLevMatrix(i).ID)
			Input(10, MagLevMatrix(i).TL)
			Input(10, MagLevMatrix(i).Weight)
			Input(10, MagLevMatrix(i).Cost)
			Input(10, MagLevMatrix(i).Volume)
			Input(10, MagLevMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadStarDrive()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\1017.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve StarDriveMatrix(i)
			Input(10, StarDriveMatrix(i).ID)
			Input(10, StarDriveMatrix(i).TL)
			Input(10, StarDriveMatrix(i).Weight1)
			Input(10, StarDriveMatrix(i).Weight2)
			Input(10, StarDriveMatrix(i).Cost1)
			Input(10, StarDriveMatrix(i).Cost2)
			Input(10, StarDriveMatrix(i).Volume)
			Input(10, StarDriveMatrix(i).Power1)
			Input(10, StarDriveMatrix(i).Power2) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadLiftingGas()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\1018.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve LiftingGasMatrix(i)
			Input(10, LiftingGasMatrix(i).ID)
			Input(10, LiftingGasMatrix(i).TL)
			Input(10, LiftingGasMatrix(i).Cost)
			Input(10, LiftingGasMatrix(i).Volume)
			Input(10, LiftingGasMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadContraGravity()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\1019.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve ContraGravMatrix(i)
			Input(10, ContraGravMatrix(i).ID)
			Input(10, ContraGravMatrix(i).TL)
			Input(10, ContraGravMatrix(i).Weight1)
			Input(10, ContraGravMatrix(i).Weight2)
			Input(10, ContraGravMatrix(i).Cost1)
			Input(10, ContraGravMatrix(i).Cost2)
			Input(10, ContraGravMatrix(i).Volume)
			Input(10, ContraGravMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
End Module