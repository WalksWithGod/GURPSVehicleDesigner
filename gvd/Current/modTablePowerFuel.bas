Attribute VB_Name = "modTablePowerFuel"
Option Explicit


Public Type udtMuscleEngine
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    MaxOutput As Single
    Output As Single
End Type

Public Type udtSteamEngine
    ID As Integer
    TL As Integer
    Weight1 As Single
    Weight2 As Single
    Weight3 As Single
    Volume As Single
    Cost As Single
    FuelUsed As Single
    FuelType As Integer
    FuelUsed2 As Single
    FuelType2 As Integer
    ElementalEnhancedMod As Single
End Type

Public Type udtCombustionEngine
    ID As Integer
    TL As Integer
    Weight1 As Single
    Weight2 As Single
    Weight3 As Single
    Volume As Single
    Cost As Single
    FuelUsed As Single
    FuelType As Integer
    AlcoholModifier As Single
    AlcoholWeight As Single
    AlcoholCost As Single
    AlcoholVolume As Single
    AlcoholFuelMod As Single
    AlcoholPower As Single
    PropaneWeight As Single
    PropaneCost As Single
    PropaneVolume As Single
    PropaneFuelMod As Single
End Type

Public Type udtNitrousOxideBooster
    ID As Integer
    TL As Integer
    PowerBoost As Single
    Weight As Single
    Volume As Single
    Cost As Single
End Type

Public Type udtSnorkel
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    SizeMod As Integer
End Type

Public Type udtTurbine
    ID As Integer
    TL As Integer
    Weight1 As Single
    Weight2 As Single
    Weight3 As Single
    Volume As Single
    Cost As Single
    MinCost As Single
    FuelUsed As Single
    CCWeight As Single
    CCVolume As Single
    CCCost As Single
    CCFuel As Single
    FuelType As Integer
    AlcoholMod As Single
End Type

Public Type udtFuelCell
    ID As Integer
    TL As Integer
    Weight1 As Single
    Weight2 As Single
    Weight3 As Single
    Cost As Single
    Volume As Single
    FuelUsed As Single
    MinCost As Single
    CCFuel As Single
    FuelType As Integer
End Type
    
Public Type udtReactor
    ID As Integer
    TL As Integer
    Weight1 As Single
    Weight2 As Single
    Weight3 As Single
    Cost As Single
    Volume As Single
    years As Single
    MinPower As Single
    CoreCost As Single
    MinCost As Single
    FuelCost As Single
    FuelNeeded As Single
End Type

Public Type udtExoticEngine
    ID As Integer
    TL As Integer
    Weight1 As Single
    Weight2 As Single
    Weight3 As Single
    Cost As Single
    Volume As Single
    MinCost As Single
    MinPower As Single
    MagicMod As Single
End Type
    
Public Type udtElectricContact
    ID As Integer
    TL As Integer
    Weight As Single
    Cost As Single
    Volume As Single
End Type
    
Public Type udtBeamedPower
    ID As Integer
    TL As Integer
    Weight As Single
    Cost As Single
    Volume As Single
    BaseCost As Single
End Type

Public Type udtEnergyBank
    ID As Integer
    TL As Integer
    Weight As Single
    Cost As Single
    Volume As Single
    PoweredClockCost As Single
    EffectiveST As Single
End Type

Public Type udtFuelTank
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Fire As Integer
End Type

Public Type udtFuel
    ID As Integer
    TL As Integer
    Weight As Single
    Cost As Single
    Fire As Integer
End Type


Public MuscleEngineMatrix() As udtMuscleEngine
Public SteamEngineMatrix() As udtSteamEngine
Public CombustionEngineMatrix() As udtCombustionEngine
Public NitrousOxideBoosterMatrix() As udtNitrousOxideBooster
Public SnorkelMatrix() As udtSnorkel
Public TurbineMatrix() As udtTurbine
Public FuelCellMatrix() As udtFuelCell
Public ReactorMatrix() As udtReactor
Public ExoticEngineMatrix() As udtExoticEngine
Public ElectricContactMatrix() As udtElectricContact
Public BeamedPowerMatrix() As udtBeamedPower
Public EnergyBankMatrix() As udtEnergyBank
Public FuelTankMatrix() As udtFuelTank
Public FuelMatrix() As udtFuel

Sub LoadMuscleEngines()
Dim i As Integer
Open App.Path & "\data\6001.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve MuscleEngineMatrix(i)
            Input #10, MuscleEngineMatrix(i).ID, MuscleEngineMatrix(i).TL, MuscleEngineMatrix(i).Weight, MuscleEngineMatrix(i).Volume, MuscleEngineMatrix(i).Cost, MuscleEngineMatrix(i).MaxOutput, MuscleEngineMatrix(i).Output  ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadSteamEngines()
Dim i As Integer
Open App.Path & "\data\6002.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve SteamEngineMatrix(i)
            Input #10, SteamEngineMatrix(i).ID, SteamEngineMatrix(i).TL, SteamEngineMatrix(i).Weight1, SteamEngineMatrix(i).Weight2, SteamEngineMatrix(i).Weight3, SteamEngineMatrix(i).Volume, SteamEngineMatrix(i).Cost, SteamEngineMatrix(i).FuelUsed, SteamEngineMatrix(i).FuelType, SteamEngineMatrix(i).FuelUsed2, SteamEngineMatrix(i).FuelType2, SteamEngineMatrix(i).ElementalEnhancedMod   ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub


Sub LoadCombustionEngines()
Dim i As Integer
Open App.Path & "\data\6003.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve CombustionEngineMatrix(i)
            Input #10, CombustionEngineMatrix(i).ID, CombustionEngineMatrix(i).TL, CombustionEngineMatrix(i).Weight1, CombustionEngineMatrix(i).Weight2, CombustionEngineMatrix(i).Weight3, CombustionEngineMatrix(i).Volume, CombustionEngineMatrix(i).Cost, CombustionEngineMatrix(i).FuelUsed, CombustionEngineMatrix(i).FuelType, CombustionEngineMatrix(i).AlcoholModifier, CombustionEngineMatrix(i).AlcoholWeight, CombustionEngineMatrix(i).AlcoholCost, CombustionEngineMatrix(i).AlcoholVolume, CombustionEngineMatrix(i).AlcoholFuelMod, CombustionEngineMatrix(i).AlcoholPower, CombustionEngineMatrix(i).PropaneWeight, CombustionEngineMatrix(i).PropaneCost, CombustionEngineMatrix(i).PropaneVolume, CombustionEngineMatrix(i).PropaneFuelMod        ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub


Sub LoadNitrousOxideBoosters()
Dim i As Integer
Open App.Path & "\data\6004.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve NitrousOxideBoosterMatrix(i)
            Input #10, NitrousOxideBoosterMatrix(i).ID, NitrousOxideBoosterMatrix(i).TL, NitrousOxideBoosterMatrix(i).PowerBoost, NitrousOxideBoosterMatrix(i).Weight, NitrousOxideBoosterMatrix(i).Volume, NitrousOxideBoosterMatrix(i).Cost ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub


Sub LoadSnorkels()
Dim i As Integer
Open App.Path & "\data\6005.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve SnorkelMatrix(i)
            Input #10, SnorkelMatrix(i).ID, SnorkelMatrix(i).TL, SnorkelMatrix(i).Weight, SnorkelMatrix(i).Volume, SnorkelMatrix(i).Cost, SnorkelMatrix(i).SizeMod  ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub


Sub LoadTurbines()
Dim i As Integer
Open App.Path & "\data\6006.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve TurbineMatrix(i)
            Input #10, TurbineMatrix(i).ID, TurbineMatrix(i).TL, TurbineMatrix(i).Weight1, TurbineMatrix(i).Weight2, TurbineMatrix(i).Weight3, TurbineMatrix(i).Volume, TurbineMatrix(i).Cost, TurbineMatrix(i).MinCost, TurbineMatrix(i).FuelUsed, TurbineMatrix(i).CCWeight, TurbineMatrix(i).CCVolume, TurbineMatrix(i).CCCost, TurbineMatrix(i).CCFuel, TurbineMatrix(i).FuelType, TurbineMatrix(i).AlcoholMod ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadFuelCells()
Dim i As Integer
Open App.Path & "\data\6007.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve FuelCellMatrix(i)
            Input #10, FuelCellMatrix(i).ID, FuelCellMatrix(i).TL, FuelCellMatrix(i).Weight1, FuelCellMatrix(i).Weight2, FuelCellMatrix(i).Weight3, FuelCellMatrix(i).Cost, FuelCellMatrix(i).Volume, FuelCellMatrix(i).FuelUsed, FuelCellMatrix(i).MinCost, FuelCellMatrix(i).CCFuel, FuelCellMatrix(i).FuelType ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadReactors()
Dim i As Integer
Open App.Path & "\data\6008.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve ReactorMatrix(i)
            Input #10, ReactorMatrix(i).ID, ReactorMatrix(i).TL, ReactorMatrix(i).Weight1, ReactorMatrix(i).Weight2, ReactorMatrix(i).Weight3, ReactorMatrix(i).Cost, ReactorMatrix(i).Volume, ReactorMatrix(i).years, ReactorMatrix(i).MinPower, ReactorMatrix(i).CoreCost, ReactorMatrix(i).MinCost, ReactorMatrix(i).FuelCost, ReactorMatrix(i).FuelNeeded  ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadExoticEngines()
Dim i As Integer
Open App.Path & "\data\6009.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve ExoticEngineMatrix(i)
            Input #10, ExoticEngineMatrix(i).ID, ExoticEngineMatrix(i).TL, ExoticEngineMatrix(i).Weight1, ExoticEngineMatrix(i).Weight2, ExoticEngineMatrix(i).Weight3, ExoticEngineMatrix(i).Cost, ExoticEngineMatrix(i).Volume, ExoticEngineMatrix(i).MinCost, ExoticEngineMatrix(i).MinPower, ExoticEngineMatrix(i).MagicMod  ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadElectricContacts()
Dim i As Integer
Open App.Path & "\data\6010.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve ElectricContactMatrix(i)
            Input #10, ElectricContactMatrix(i).ID, ElectricContactMatrix(i).TL, ElectricContactMatrix(i).Weight, ElectricContactMatrix(i).Cost, ElectricContactMatrix(i).Volume  ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadBeamedPowers()
Dim i As Integer
Open App.Path & "\data\6011.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve BeamedPowerMatrix(i)
            Input #10, BeamedPowerMatrix(i).ID, BeamedPowerMatrix(i).TL, BeamedPowerMatrix(i).Weight, BeamedPowerMatrix(i).Cost, BeamedPowerMatrix(i).Volume, BeamedPowerMatrix(i).BaseCost   ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadEnergyBanks()
Dim i As Integer
Open App.Path & "\data\6012.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve EnergyBankMatrix(i)
            Input #10, EnergyBankMatrix(i).ID, EnergyBankMatrix(i).TL, EnergyBankMatrix(i).Weight, EnergyBankMatrix(i).Cost, EnergyBankMatrix(i).Volume, EnergyBankMatrix(i).PoweredClockCost, EnergyBankMatrix(i).EffectiveST     ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadFuelTanks()
Dim i As Integer
Open App.Path & "\data\6013.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve FuelTankMatrix(i)
            Input #10, FuelTankMatrix(i).ID, FuelTankMatrix(i).TL, FuelTankMatrix(i).Weight, FuelTankMatrix(i).Volume, FuelTankMatrix(i).Cost, FuelTankMatrix(i).Fire     ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadFuel()
Dim i As Integer
Open App.Path & "\data\6014.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve FuelMatrix(i)
            Input #10, FuelMatrix(i).ID, FuelMatrix(i).TL, FuelMatrix(i).Weight, FuelMatrix(i).Cost, FuelMatrix(i).Fire     ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub



