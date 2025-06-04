Attribute VB_Name = "modTableInstrumentsElectronics"
Option Explicit

Public Type udtCommunicator
    ID As Integer
    TL As Integer
    Weight As Single
    Cost As Single
    Volume As Single
    Range As Single
    Power As Single
End Type

Public Type udtScramblerOptions
    TL6Cost As Single
    TL7Cost As Single
    TL8Cost As Single
    TL9Cost As Single
    TL10Cost As Single
End Type

Public Type udtRadioOptions
    Weight As Single
    Cost As Single
    Range As Single
    Power As Single
End Type

Public Type udtSearchlight
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtVisualAugmentation
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type
    
Public Type udtRadar
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtSonar
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtSonarOptions
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtThermalPassiveIR
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtOtherSensor
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtSoundDetector
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtScientificSensor
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtAVSystem
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtNavigationSystem
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtTargetingSystem
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtCounterMeasure
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtDecoyReload
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
End Type

Public Type udtComputer
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
    Complexity As Integer
End Type

Public Type udtTerminal
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtSoftware
    ID As Integer
    TL As Integer
    Cost As Single
    Complexity As Integer
    BonusSkill As Integer
    Skill As Integer
End Type

Public Type udtNeuralInterface
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtShields
    ID As Integer
    TL As Integer
    Weight As Single
    Cost As Single
    Power As Single
End Type

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
Dim i As Integer
Open App.Path & "\data\3001.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve CommunicatorMatrix(i)
            Input #10, CommunicatorMatrix(i).ID, CommunicatorMatrix(i).TL, CommunicatorMatrix(i).Weight, CommunicatorMatrix(i).Cost, CommunicatorMatrix(i).Volume, CommunicatorMatrix(i).Range, CommunicatorMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadScramblerOptions()
Dim i As Integer
Open App.Path & "\data\3002.txt" For Input As #10
i = 1
Do While Not EOF(10)
    ReDim Preserve ScramblerOptionsMatrix(i)
        Input #10, ScramblerOptionsMatrix(i).TL6Cost, ScramblerOptionsMatrix(i).TL7Cost, ScramblerOptionsMatrix(i).TL8Cost, ScramblerOptionsMatrix(i).TL9Cost, ScramblerOptionsMatrix(i).TL10Cost
    i = i + 1
Loop
Close #10
End Sub

Sub LoadRadioOptions()
Dim i As Integer
Open App.Path & "\data\3003.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve RadioOptionsMatrix(i)
            Input #10, RadioOptionsMatrix(i).Weight, RadioOptionsMatrix(i).Cost, RadioOptionsMatrix(i).Range, RadioOptionsMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadSearchlights()
Dim i As Integer
Open App.Path & "\data\3004.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve SearchlightMatrix(i)
            Input #10, SearchlightMatrix(i).ID, SearchlightMatrix(i).TL, SearchlightMatrix(i).Weight, SearchlightMatrix(i).Volume, SearchlightMatrix(i).Cost, SearchlightMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadVisualAugmentations()
Dim i As Integer
Open App.Path & "\data\3005.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve VisualAugmentationMatrix(i)
            Input #10, VisualAugmentationMatrix(i).ID, VisualAugmentationMatrix(i).TL, VisualAugmentationMatrix(i).Weight, VisualAugmentationMatrix(i).Volume, VisualAugmentationMatrix(i).Cost, VisualAugmentationMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadRadars()
Dim i As Integer
Open App.Path & "\data\3006.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve RadarMatrix(i)
            Input #10, RadarMatrix(i).ID, RadarMatrix(i).TL, RadarMatrix(i).Weight, RadarMatrix(i).Volume, RadarMatrix(i).Cost, RadarMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadSonars()
Dim i As Integer
Open App.Path & "\data\3007.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve SonarMatrix(i)
            Input #10, SonarMatrix(i).ID, SonarMatrix(i).TL, SonarMatrix(i).Weight, SonarMatrix(i).Volume, SonarMatrix(i).Cost, SonarMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadSonarOptions()
Dim i As Integer
Open App.Path & "\data\3008.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve SonarOptionsMatrix(i)
            Input #10, SonarOptionsMatrix(i).Weight, SonarOptionsMatrix(i).Volume, SonarOptionsMatrix(i).Cost, SonarOptionsMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadThermalPassiveIRs()
Dim i As Integer
Open App.Path & "\data\3009.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve ThermalPassiveIRMatrix(i)
            Input #10, ThermalPassiveIRMatrix(i).ID, ThermalPassiveIRMatrix(i).TL, ThermalPassiveIRMatrix(i).Weight, ThermalPassiveIRMatrix(i).Volume, ThermalPassiveIRMatrix(i).Cost, ThermalPassiveIRMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadOtherSensors()
Dim i As Integer
Open App.Path & "\data\3010.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve OtherSensorMatrix(i)
            Input #10, OtherSensorMatrix(i).ID, OtherSensorMatrix(i).TL, OtherSensorMatrix(i).Weight, OtherSensorMatrix(i).Volume, OtherSensorMatrix(i).Cost, OtherSensorMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadSoundDetectors()
Dim i As Integer
Open App.Path & "\data\3011.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve SoundDetectorMatrix(i)
            Input #10, SoundDetectorMatrix(i).ID, SoundDetectorMatrix(i).TL, SoundDetectorMatrix(i).Weight, SoundDetectorMatrix(i).Volume, SoundDetectorMatrix(i).Cost, SoundDetectorMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub


Sub LoadScientificSensors()
Dim i As Integer
Open App.Path & "\data\3012.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve ScientificSensorMatrix(i)
            Input #10, ScientificSensorMatrix(i).ID, ScientificSensorMatrix(i).TL, ScientificSensorMatrix(i).Weight, ScientificSensorMatrix(i).Volume, ScientificSensorMatrix(i).Cost, ScientificSensorMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadAVSystems()
Dim i As Integer
Open App.Path & "\data\3013.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve AVSystemMatrix(i)
            Input #10, AVSystemMatrix(i).ID, AVSystemMatrix(i).TL, AVSystemMatrix(i).Weight, AVSystemMatrix(i).Volume, AVSystemMatrix(i).Cost, AVSystemMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadNavigationSystems()
Dim i As Integer
Open App.Path & "\data\3014.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve NavigationSystemMatrix(i)
            Input #10, NavigationSystemMatrix(i).ID, NavigationSystemMatrix(i).TL, NavigationSystemMatrix(i).Weight, NavigationSystemMatrix(i).Volume, NavigationSystemMatrix(i).Cost, NavigationSystemMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadTargetingSystems()
Dim i As Integer
Open App.Path & "\data\3015.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve TargetingSystemMatrix(i)
            Input #10, TargetingSystemMatrix(i).ID, TargetingSystemMatrix(i).TL, TargetingSystemMatrix(i).Weight, TargetingSystemMatrix(i).Volume, TargetingSystemMatrix(i).Cost, TargetingSystemMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadCounterMeasures()
Dim i As Integer
Open App.Path & "\data\3016.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve CounterMeasureMatrix(i)
            Input #10, CounterMeasureMatrix(i).ID, CounterMeasureMatrix(i).TL, CounterMeasureMatrix(i).Weight, CounterMeasureMatrix(i).Volume, CounterMeasureMatrix(i).Cost, CounterMeasureMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub


Sub LoadComputers()
Dim i As Integer
Open App.Path & "\data\3017.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve ComputerMatrix(i)
            Input #10, ComputerMatrix(i).ID, ComputerMatrix(i).TL, ComputerMatrix(i).Weight, ComputerMatrix(i).Volume, ComputerMatrix(i).Cost, ComputerMatrix(i).Power, ComputerMatrix(i).Complexity  ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadTerminals()
Dim i As Integer
Open App.Path & "\data\3018.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve TerminalMatrix(i)
            Input #10, TerminalMatrix(i).ID, TerminalMatrix(i).TL, TerminalMatrix(i).Weight, TerminalMatrix(i).Volume, TerminalMatrix(i).Cost, TerminalMatrix(i).Power  ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadSoftware()
Dim i As Integer
Open App.Path & "\data\3019.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve SoftwareMatrix(i)
            Input #10, SoftwareMatrix(i).ID, SoftwareMatrix(i).TL, SoftwareMatrix(i).Cost, SoftwareMatrix(i).Complexity, SoftwareMatrix(i).BonusSkill, SoftwareMatrix(i).Skill    ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadNeuralInterfaces()
Dim i As Integer
Open App.Path & "\data\3020.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve NeuralInterfaceMatrix(i)
            Input #10, NeuralInterfaceMatrix(i).ID, NeuralInterfaceMatrix(i).TL, NeuralInterfaceMatrix(i).Weight, NeuralInterfaceMatrix(i).Volume, NeuralInterfaceMatrix(i).Cost, NeuralInterfaceMatrix(i).Power  ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadDecoyReloads()
Dim i As Integer
Open App.Path & "\data\3021.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve DecoyReloadMatrix(i)
            Input #10, DecoyReloadMatrix(i).ID, DecoyReloadMatrix(i).TL, DecoyReloadMatrix(i).Weight, DecoyReloadMatrix(i).Volume, DecoyReloadMatrix(i).Cost
        i = i + 1
Loop
Close #10    ' Close file.
End Sub


Sub LoadShields()
Dim i As Integer
Open App.Path & "\data\3022.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve ShieldsMatrix(i)
            Input #10, ShieldsMatrix(i).ID, ShieldsMatrix(i).TL, ShieldsMatrix(i).Weight, ShieldsMatrix(i).Cost, ShieldsMatrix(i).Power   ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

