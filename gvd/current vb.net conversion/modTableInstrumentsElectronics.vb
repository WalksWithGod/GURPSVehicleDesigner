Option Strict Off
Option Explicit On
Module modTableInstrumentsElectronics
	
	Public Structure udtCommunicator
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Cost As Single
		Dim Volume As Single
		Dim Range As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtScramblerOptions
		Dim TL6Cost As Single
		Dim TL7Cost As Single
		Dim TL8Cost As Single
		Dim TL9Cost As Single
		Dim TL10Cost As Single
	End Structure
	
	Public Structure udtRadioOptions
		Dim Weight As Single
		Dim Cost As Single
		Dim Range As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtSearchlight
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtVisualAugmentation
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtRadar
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtSonar
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtSonarOptions
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtThermalPassiveIR
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtOtherSensor
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtSoundDetector
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtScientificSensor
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtAVSystem
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtNavigationSystem
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtTargetingSystem
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtCounterMeasure
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtDecoyReload
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
	End Structure
	
	Public Structure udtComputer
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
		Dim Complexity As Short
	End Structure
	
	Public Structure udtTerminal
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtSoftware
		Dim ID As Short
		Dim TL As Short
		Dim Cost As Single
		Dim Complexity As Short
		Dim BonusSkill As Short
		Dim Skill As Short
	End Structure
	
	Public Structure udtNeuralInterface
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtShields
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public CommunicatorMatrix() As udtCommunicator
	Public ScramblerOptionsMatrix() As udtScramblerOptions
	Public RadioOptionsMatrix() As udtRadioOptions
	Public SearchlightMatrix() As udtSearchlight
	Public VisualAugmentationMatrix() As udtVisualAugmentation
	Public RadarMatrix() As udtRadar
	Public SonarMatrix() As udtSonar
	Public SonarOptionsMatrix() As udtSonarOptions
	Public ThermalPassiveIRMatrix() As udtThermalPassiveIR
	Public OtherSensorMatrix() As udtSoundDetector
	Public SoundDetectorMatrix() As udtSoundDetector
	Public ScientificSensorMatrix() As udtScientificSensor
	Public AVSystemMatrix() As udtAVSystem
	Public NavigationSystemMatrix() As udtNavigationSystem
	Public TargetingSystemMatrix() As udtTargetingSystem
	Public CounterMeasureMatrix() As udtCounterMeasure
	Public DecoyReloadMatrix() As udtDecoyReload
	Public ComputerMatrix() As udtComputer
	Public TerminalMatrix() As udtTerminal
	Public SoftwareMatrix() As udtSoftware
	Public NeuralInterfaceMatrix() As udtNeuralInterface
	Public ShieldsMatrix() As udtShields
	
	Sub LoadCommunicators()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\3001.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve CommunicatorMatrix(i)
			Input(10, CommunicatorMatrix(i).ID)
			Input(10, CommunicatorMatrix(i).TL)
			Input(10, CommunicatorMatrix(i).Weight)
			Input(10, CommunicatorMatrix(i).Cost)
			Input(10, CommunicatorMatrix(i).Volume)
			Input(10, CommunicatorMatrix(i).Range)
			Input(10, CommunicatorMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadScramblerOptions()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\3002.txt", OpenMode.Input)
		i = 1
		Do While Not EOF(10)
			ReDim Preserve ScramblerOptionsMatrix(i)
			Input(10, ScramblerOptionsMatrix(i).TL6Cost)
			Input(10, ScramblerOptionsMatrix(i).TL7Cost)
			Input(10, ScramblerOptionsMatrix(i).TL8Cost)
			Input(10, ScramblerOptionsMatrix(i).TL9Cost)
			Input(10, ScramblerOptionsMatrix(i).TL10Cost)
			i = i + 1
		Loop 
		FileClose(10)
	End Sub
	
	Sub LoadRadioOptions()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\3003.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve RadioOptionsMatrix(i)
			Input(10, RadioOptionsMatrix(i).Weight)
			Input(10, RadioOptionsMatrix(i).Cost)
			Input(10, RadioOptionsMatrix(i).Range)
			Input(10, RadioOptionsMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadSearchlights()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\3004.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve SearchlightMatrix(i)
			Input(10, SearchlightMatrix(i).ID)
			Input(10, SearchlightMatrix(i).TL)
			Input(10, SearchlightMatrix(i).Weight)
			Input(10, SearchlightMatrix(i).Volume)
			Input(10, SearchlightMatrix(i).Cost)
			Input(10, SearchlightMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadVisualAugmentations()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\3005.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve VisualAugmentationMatrix(i)
			Input(10, VisualAugmentationMatrix(i).ID)
			Input(10, VisualAugmentationMatrix(i).TL)
			Input(10, VisualAugmentationMatrix(i).Weight)
			Input(10, VisualAugmentationMatrix(i).Volume)
			Input(10, VisualAugmentationMatrix(i).Cost)
			Input(10, VisualAugmentationMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadRadars()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\3006.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve RadarMatrix(i)
			Input(10, RadarMatrix(i).ID)
			Input(10, RadarMatrix(i).TL)
			Input(10, RadarMatrix(i).Weight)
			Input(10, RadarMatrix(i).Volume)
			Input(10, RadarMatrix(i).Cost)
			Input(10, RadarMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadSonars()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\3007.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve SonarMatrix(i)
			Input(10, SonarMatrix(i).ID)
			Input(10, SonarMatrix(i).TL)
			Input(10, SonarMatrix(i).Weight)
			Input(10, SonarMatrix(i).Volume)
			Input(10, SonarMatrix(i).Cost)
			Input(10, SonarMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadSonarOptions()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\3008.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve SonarOptionsMatrix(i)
			Input(10, SonarOptionsMatrix(i).Weight)
			Input(10, SonarOptionsMatrix(i).Volume)
			Input(10, SonarOptionsMatrix(i).Cost)
			Input(10, SonarOptionsMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadThermalPassiveIRs()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\3009.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve ThermalPassiveIRMatrix(i)
			Input(10, ThermalPassiveIRMatrix(i).ID)
			Input(10, ThermalPassiveIRMatrix(i).TL)
			Input(10, ThermalPassiveIRMatrix(i).Weight)
			Input(10, ThermalPassiveIRMatrix(i).Volume)
			Input(10, ThermalPassiveIRMatrix(i).Cost)
			Input(10, ThermalPassiveIRMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadOtherSensors()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\3010.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve OtherSensorMatrix(i)
			Input(10, OtherSensorMatrix(i).ID)
			Input(10, OtherSensorMatrix(i).TL)
			Input(10, OtherSensorMatrix(i).Weight)
			Input(10, OtherSensorMatrix(i).Volume)
			Input(10, OtherSensorMatrix(i).Cost)
			Input(10, OtherSensorMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadSoundDetectors()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\3011.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve SoundDetectorMatrix(i)
			Input(10, SoundDetectorMatrix(i).ID)
			Input(10, SoundDetectorMatrix(i).TL)
			Input(10, SoundDetectorMatrix(i).Weight)
			Input(10, SoundDetectorMatrix(i).Volume)
			Input(10, SoundDetectorMatrix(i).Cost)
			Input(10, SoundDetectorMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	
	Sub LoadScientificSensors()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\3012.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve ScientificSensorMatrix(i)
			Input(10, ScientificSensorMatrix(i).ID)
			Input(10, ScientificSensorMatrix(i).TL)
			Input(10, ScientificSensorMatrix(i).Weight)
			Input(10, ScientificSensorMatrix(i).Volume)
			Input(10, ScientificSensorMatrix(i).Cost)
			Input(10, ScientificSensorMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadAVSystems()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\3013.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve AVSystemMatrix(i)
			Input(10, AVSystemMatrix(i).ID)
			Input(10, AVSystemMatrix(i).TL)
			Input(10, AVSystemMatrix(i).Weight)
			Input(10, AVSystemMatrix(i).Volume)
			Input(10, AVSystemMatrix(i).Cost)
			Input(10, AVSystemMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadNavigationSystems()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\3014.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve NavigationSystemMatrix(i)
			Input(10, NavigationSystemMatrix(i).ID)
			Input(10, NavigationSystemMatrix(i).TL)
			Input(10, NavigationSystemMatrix(i).Weight)
			Input(10, NavigationSystemMatrix(i).Volume)
			Input(10, NavigationSystemMatrix(i).Cost)
			Input(10, NavigationSystemMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadTargetingSystems()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\3015.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve TargetingSystemMatrix(i)
			Input(10, TargetingSystemMatrix(i).ID)
			Input(10, TargetingSystemMatrix(i).TL)
			Input(10, TargetingSystemMatrix(i).Weight)
			Input(10, TargetingSystemMatrix(i).Volume)
			Input(10, TargetingSystemMatrix(i).Cost)
			Input(10, TargetingSystemMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadCounterMeasures()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\3016.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve CounterMeasureMatrix(i)
			Input(10, CounterMeasureMatrix(i).ID)
			Input(10, CounterMeasureMatrix(i).TL)
			Input(10, CounterMeasureMatrix(i).Weight)
			Input(10, CounterMeasureMatrix(i).Volume)
			Input(10, CounterMeasureMatrix(i).Cost)
			Input(10, CounterMeasureMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	
	Sub LoadComputers()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\3017.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve ComputerMatrix(i)
			Input(10, ComputerMatrix(i).ID)
			Input(10, ComputerMatrix(i).TL)
			Input(10, ComputerMatrix(i).Weight)
			Input(10, ComputerMatrix(i).Volume)
			Input(10, ComputerMatrix(i).Cost)
			Input(10, ComputerMatrix(i).Power)
			Input(10, ComputerMatrix(i).Complexity) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadTerminals()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\3018.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve TerminalMatrix(i)
			Input(10, TerminalMatrix(i).ID)
			Input(10, TerminalMatrix(i).TL)
			Input(10, TerminalMatrix(i).Weight)
			Input(10, TerminalMatrix(i).Volume)
			Input(10, TerminalMatrix(i).Cost)
			Input(10, TerminalMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadSoftware()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\3019.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve SoftwareMatrix(i)
			Input(10, SoftwareMatrix(i).ID)
			Input(10, SoftwareMatrix(i).TL)
			Input(10, SoftwareMatrix(i).Cost)
			Input(10, SoftwareMatrix(i).Complexity)
			Input(10, SoftwareMatrix(i).BonusSkill)
			Input(10, SoftwareMatrix(i).Skill) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadNeuralInterfaces()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\3020.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve NeuralInterfaceMatrix(i)
			Input(10, NeuralInterfaceMatrix(i).ID)
			Input(10, NeuralInterfaceMatrix(i).TL)
			Input(10, NeuralInterfaceMatrix(i).Weight)
			Input(10, NeuralInterfaceMatrix(i).Volume)
			Input(10, NeuralInterfaceMatrix(i).Cost)
			Input(10, NeuralInterfaceMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadDecoyReloads()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\3021.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve DecoyReloadMatrix(i)
			Input(10, DecoyReloadMatrix(i).ID)
			Input(10, DecoyReloadMatrix(i).TL)
			Input(10, DecoyReloadMatrix(i).Weight)
			Input(10, DecoyReloadMatrix(i).Volume)
			Input(10, DecoyReloadMatrix(i).Cost)
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	
	Sub LoadShields()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\3022.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve ShieldsMatrix(i)
			Input(10, ShieldsMatrix(i).ID)
			Input(10, ShieldsMatrix(i).TL)
			Input(10, ShieldsMatrix(i).Weight)
			Input(10, ShieldsMatrix(i).Cost)
			Input(10, ShieldsMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
End Module