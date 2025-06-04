Option Strict Off
Option Explicit On
Module modTablePowerFuel
	
	
	Public Structure udtMuscleEngine
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim MaxOutput As Single
		Dim Output As Single
	End Structure
	
	Public Structure udtSteamEngine
		Dim ID As Short
		Dim TL As Short
		Dim Weight1 As Single
		Dim Weight2 As Single
		Dim Weight3 As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim FuelUsed As Single
		Dim FuelType As Short
		Dim FuelUsed2 As Single
		Dim FuelType2 As Short
		Dim ElementalEnhancedMod As Single
	End Structure
	
	Public Structure udtCombustionEngine
		Dim ID As Short
		Dim TL As Short
		Dim Weight1 As Single
		Dim Weight2 As Single
		Dim Weight3 As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim FuelUsed As Single
		Dim FuelType As Short
		Dim AlcoholModifier As Single
		Dim AlcoholWeight As Single
		Dim AlcoholCost As Single
		Dim AlcoholVolume As Single
		Dim AlcoholFuelMod As Single
		Dim AlcoholPower As Single
		Dim PropaneWeight As Single
		Dim PropaneCost As Single
		Dim PropaneVolume As Single
		Dim PropaneFuelMod As Single
	End Structure
	
	Public Structure udtNitrousOxideBooster
		Dim ID As Short
		Dim TL As Short
		Dim PowerBoost As Single
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
	End Structure
	
	Public Structure udtSnorkel
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim SizeMod As Short
	End Structure
	
	Public Structure udtTurbine
		Dim ID As Short
		Dim TL As Short
		Dim Weight1 As Single
		Dim Weight2 As Single
		Dim Weight3 As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim MinCost As Single
		Dim FuelUsed As Single
		Dim CCWeight As Single
		Dim CCVolume As Single
		Dim CCCost As Single
		Dim CCFuel As Single
		Dim FuelType As Short
		Dim AlcoholMod As Single
	End Structure
	
	Public Structure udtFuelCell
		Dim ID As Short
		Dim TL As Short
		Dim Weight1 As Single
		Dim Weight2 As Single
		Dim Weight3 As Single
		Dim Cost As Single
		Dim Volume As Single
		Dim FuelUsed As Single
		Dim MinCost As Single
		Dim CCFuel As Single
		Dim FuelType As Short
	End Structure
	
	Public Structure udtReactor
		Dim ID As Short
		Dim TL As Short
		Dim Weight1 As Single
		Dim Weight2 As Single
		Dim Weight3 As Single
		Dim Cost As Single
		Dim Volume As Single
		Dim years As Single
		Dim MinPower As Single
		Dim CoreCost As Single
		Dim MinCost As Single
		Dim FuelCost As Single
		Dim FuelNeeded As Single
	End Structure
	
	Public Structure udtExoticEngine
		Dim ID As Short
		Dim TL As Short
		Dim Weight1 As Single
		Dim Weight2 As Single
		Dim Weight3 As Single
		Dim Cost As Single
		Dim Volume As Single
		Dim MinCost As Single
		Dim MinPower As Single
		Dim MagicMod As Single
	End Structure
	
	Public Structure udtElectricContact
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Cost As Single
		Dim Volume As Single
	End Structure
	
	Public Structure udtBeamedPower
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Cost As Single
		Dim Volume As Single
		Dim BaseCost As Single
	End Structure
	
	Public Structure udtEnergyBank
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Cost As Single
		Dim Volume As Single
		Dim PoweredClockCost As Single
		Dim EffectiveST As Single
	End Structure
	
	Public Structure udtFuelTank
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Fire As Short
	End Structure
	
	Public Structure udtFuel
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Cost As Single
		Dim Fire As Short
	End Structure
	
	
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
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\6001.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve MuscleEngineMatrix(i)
			Input(10, MuscleEngineMatrix(i).ID)
			Input(10, MuscleEngineMatrix(i).TL)
			Input(10, MuscleEngineMatrix(i).Weight)
			Input(10, MuscleEngineMatrix(i).Volume)
			Input(10, MuscleEngineMatrix(i).Cost)
			Input(10, MuscleEngineMatrix(i).MaxOutput)
			Input(10, MuscleEngineMatrix(i).Output) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadSteamEngines()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\6002.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve SteamEngineMatrix(i)
			Input(10, SteamEngineMatrix(i).ID)
			Input(10, SteamEngineMatrix(i).TL)
			Input(10, SteamEngineMatrix(i).Weight1)
			Input(10, SteamEngineMatrix(i).Weight2)
			Input(10, SteamEngineMatrix(i).Weight3)
			Input(10, SteamEngineMatrix(i).Volume)
			Input(10, SteamEngineMatrix(i).Cost)
			Input(10, SteamEngineMatrix(i).FuelUsed)
			Input(10, SteamEngineMatrix(i).FuelType)
			Input(10, SteamEngineMatrix(i).FuelUsed2)
			Input(10, SteamEngineMatrix(i).FuelType2)
			Input(10, SteamEngineMatrix(i).ElementalEnhancedMod) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	
	Sub LoadCombustionEngines()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\6003.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve CombustionEngineMatrix(i)
			Input(10, CombustionEngineMatrix(i).ID)
			Input(10, CombustionEngineMatrix(i).TL)
			Input(10, CombustionEngineMatrix(i).Weight1)
			Input(10, CombustionEngineMatrix(i).Weight2)
			Input(10, CombustionEngineMatrix(i).Weight3)
			Input(10, CombustionEngineMatrix(i).Volume)
			Input(10, CombustionEngineMatrix(i).Cost)
			Input(10, CombustionEngineMatrix(i).FuelUsed)
			Input(10, CombustionEngineMatrix(i).FuelType)
			Input(10, CombustionEngineMatrix(i).AlcoholModifier)
			Input(10, CombustionEngineMatrix(i).AlcoholWeight)
			Input(10, CombustionEngineMatrix(i).AlcoholCost)
			Input(10, CombustionEngineMatrix(i).AlcoholVolume)
			Input(10, CombustionEngineMatrix(i).AlcoholFuelMod)
			Input(10, CombustionEngineMatrix(i).AlcoholPower)
			Input(10, CombustionEngineMatrix(i).PropaneWeight)
			Input(10, CombustionEngineMatrix(i).PropaneCost)
			Input(10, CombustionEngineMatrix(i).PropaneVolume)
			Input(10, CombustionEngineMatrix(i).PropaneFuelMod) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	
	Sub LoadNitrousOxideBoosters()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\6004.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve NitrousOxideBoosterMatrix(i)
			Input(10, NitrousOxideBoosterMatrix(i).ID)
			Input(10, NitrousOxideBoosterMatrix(i).TL)
			Input(10, NitrousOxideBoosterMatrix(i).PowerBoost)
			Input(10, NitrousOxideBoosterMatrix(i).Weight)
			Input(10, NitrousOxideBoosterMatrix(i).Volume)
			Input(10, NitrousOxideBoosterMatrix(i).Cost) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	
	Sub LoadSnorkels()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\6005.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve SnorkelMatrix(i)
			Input(10, SnorkelMatrix(i).ID)
			Input(10, SnorkelMatrix(i).TL)
			Input(10, SnorkelMatrix(i).Weight)
			Input(10, SnorkelMatrix(i).Volume)
			Input(10, SnorkelMatrix(i).Cost)
			Input(10, SnorkelMatrix(i).SizeMod) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	
	Sub LoadTurbines()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\6006.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve TurbineMatrix(i)
			Input(10, TurbineMatrix(i).ID)
			Input(10, TurbineMatrix(i).TL)
			Input(10, TurbineMatrix(i).Weight1)
			Input(10, TurbineMatrix(i).Weight2)
			Input(10, TurbineMatrix(i).Weight3)
			Input(10, TurbineMatrix(i).Volume)
			Input(10, TurbineMatrix(i).Cost)
			Input(10, TurbineMatrix(i).MinCost)
			Input(10, TurbineMatrix(i).FuelUsed)
			Input(10, TurbineMatrix(i).CCWeight)
			Input(10, TurbineMatrix(i).CCVolume)
			Input(10, TurbineMatrix(i).CCCost)
			Input(10, TurbineMatrix(i).CCFuel)
			Input(10, TurbineMatrix(i).FuelType)
			Input(10, TurbineMatrix(i).AlcoholMod) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadFuelCells()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\6007.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve FuelCellMatrix(i)
			Input(10, FuelCellMatrix(i).ID)
			Input(10, FuelCellMatrix(i).TL)
			Input(10, FuelCellMatrix(i).Weight1)
			Input(10, FuelCellMatrix(i).Weight2)
			Input(10, FuelCellMatrix(i).Weight3)
			Input(10, FuelCellMatrix(i).Cost)
			Input(10, FuelCellMatrix(i).Volume)
			Input(10, FuelCellMatrix(i).FuelUsed)
			Input(10, FuelCellMatrix(i).MinCost)
			Input(10, FuelCellMatrix(i).CCFuel)
			Input(10, FuelCellMatrix(i).FuelType) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadReactors()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\6008.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve ReactorMatrix(i)
			Input(10, ReactorMatrix(i).ID)
			Input(10, ReactorMatrix(i).TL)
			Input(10, ReactorMatrix(i).Weight1)
			Input(10, ReactorMatrix(i).Weight2)
			Input(10, ReactorMatrix(i).Weight3)
			Input(10, ReactorMatrix(i).Cost)
			Input(10, ReactorMatrix(i).Volume)
			Input(10, ReactorMatrix(i).years)
			Input(10, ReactorMatrix(i).MinPower)
			Input(10, ReactorMatrix(i).CoreCost)
			Input(10, ReactorMatrix(i).MinCost)
			Input(10, ReactorMatrix(i).FuelCost)
			Input(10, ReactorMatrix(i).FuelNeeded) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadExoticEngines()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\6009.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve ExoticEngineMatrix(i)
			Input(10, ExoticEngineMatrix(i).ID)
			Input(10, ExoticEngineMatrix(i).TL)
			Input(10, ExoticEngineMatrix(i).Weight1)
			Input(10, ExoticEngineMatrix(i).Weight2)
			Input(10, ExoticEngineMatrix(i).Weight3)
			Input(10, ExoticEngineMatrix(i).Cost)
			Input(10, ExoticEngineMatrix(i).Volume)
			Input(10, ExoticEngineMatrix(i).MinCost)
			Input(10, ExoticEngineMatrix(i).MinPower)
			Input(10, ExoticEngineMatrix(i).MagicMod) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadElectricContacts()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\6010.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve ElectricContactMatrix(i)
			Input(10, ElectricContactMatrix(i).ID)
			Input(10, ElectricContactMatrix(i).TL)
			Input(10, ElectricContactMatrix(i).Weight)
			Input(10, ElectricContactMatrix(i).Cost)
			Input(10, ElectricContactMatrix(i).Volume) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadBeamedPowers()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\6011.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve BeamedPowerMatrix(i)
			Input(10, BeamedPowerMatrix(i).ID)
			Input(10, BeamedPowerMatrix(i).TL)
			Input(10, BeamedPowerMatrix(i).Weight)
			Input(10, BeamedPowerMatrix(i).Cost)
			Input(10, BeamedPowerMatrix(i).Volume)
			Input(10, BeamedPowerMatrix(i).BaseCost) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadEnergyBanks()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\6012.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve EnergyBankMatrix(i)
			Input(10, EnergyBankMatrix(i).ID)
			Input(10, EnergyBankMatrix(i).TL)
			Input(10, EnergyBankMatrix(i).Weight)
			Input(10, EnergyBankMatrix(i).Cost)
			Input(10, EnergyBankMatrix(i).Volume)
			Input(10, EnergyBankMatrix(i).PoweredClockCost)
			Input(10, EnergyBankMatrix(i).EffectiveST) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadFuelTanks()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\6013.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve FuelTankMatrix(i)
			Input(10, FuelTankMatrix(i).ID)
			Input(10, FuelTankMatrix(i).TL)
			Input(10, FuelTankMatrix(i).Weight)
			Input(10, FuelTankMatrix(i).Volume)
			Input(10, FuelTankMatrix(i).Cost)
			Input(10, FuelTankMatrix(i).Fire) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadFuel()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\6014.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve FuelMatrix(i)
			Input(10, FuelMatrix(i).ID)
			Input(10, FuelMatrix(i).TL)
			Input(10, FuelMatrix(i).Weight)
			Input(10, FuelMatrix(i).Cost)
			Input(10, FuelMatrix(i).Fire) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
End Module