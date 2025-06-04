Option Strict Off
Option Explicit On
Module modTableMiscEquipt
	
	Public Structure udtArmMotor
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtBilgePump
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtFireExtinguisher
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
	End Structure
	
	Public Structure udtWorkshop
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
		Dim PhysicsMod As Single
		Dim ElectricMod As Single
		Dim ElectricMod2 As Single
		Dim ElectricCost As Single
	End Structure
	
	Public Structure udtHeavyequipment
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtMedical
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtEntertainment
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtVehicleAccess
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtTeleporter
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtSecurity
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtVehicleStorage
		Dim ID As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
	End Structure
	
	Public Structure udtLandingAid
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtFuelAccessory
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtNuclearDamper
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtRealityStabilizer
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtModularSocket
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	
	Public ArmMotorMatrix() As udtArmMotor
	Public BilgePumpMatrix() As udtBilgePump
	Public FireExtinguisherMatrix() As udtFireExtinguisher
	Public WorkshopMatrix() As udtWorkshop
	Public HeavyequipmentMatrix() As udtHeavyequipment
	Public MedicalMatrix() As udtMedical
	Public EntertainmentMatrix() As udtEntertainment
	Public VehicleAccessMatrix() As udtVehicleAccess
	Public TeleporterMatrix() As udtTeleporter
	Public SecurityMatrix() As udtSecurity
	Public VehicleStorageMatrix() As udtVehicleStorage
	Public LandingAidMatrix() As udtLandingAid
	Public FuelAccessoryMatrix() As udtFuelAccessory
	Public NuclearDamperMatrix() As udtNuclearDamper
	Public RealityStabilizerMatrix() As udtRealityStabilizer
	Public ModularSocketMatrix() As udtModularSocket
	
	Sub LoadArmMotors()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\4001.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve ArmMotorMatrix(i)
			Input(10, ArmMotorMatrix(i).ID)
			Input(10, ArmMotorMatrix(i).TL)
			Input(10, ArmMotorMatrix(i).Weight)
			Input(10, ArmMotorMatrix(i).Volume)
			Input(10, ArmMotorMatrix(i).Cost)
			Input(10, ArmMotorMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadBilgePumps()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\4002.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve BilgePumpMatrix(i)
			Input(10, BilgePumpMatrix(i).ID)
			Input(10, BilgePumpMatrix(i).TL)
			Input(10, BilgePumpMatrix(i).Weight)
			Input(10, BilgePumpMatrix(i).Volume)
			Input(10, BilgePumpMatrix(i).Cost)
			Input(10, BilgePumpMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadFireExtinguishers()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\4003.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve FireExtinguisherMatrix(i)
			Input(10, FireExtinguisherMatrix(i).ID)
			Input(10, FireExtinguisherMatrix(i).TL)
			Input(10, FireExtinguisherMatrix(i).Weight)
			Input(10, FireExtinguisherMatrix(i).Volume)
			Input(10, FireExtinguisherMatrix(i).Cost) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadWorkshops()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\4004.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve WorkshopMatrix(i)
			Input(10, WorkshopMatrix(i).ID)
			Input(10, WorkshopMatrix(i).TL)
			Input(10, WorkshopMatrix(i).Weight)
			Input(10, WorkshopMatrix(i).Volume)
			Input(10, WorkshopMatrix(i).Cost)
			Input(10, WorkshopMatrix(i).Power)
			Input(10, WorkshopMatrix(i).PhysicsMod)
			Input(10, WorkshopMatrix(i).ElectricMod)
			Input(10, WorkshopMatrix(i).ElectricMod2)
			Input(10, WorkshopMatrix(i).ElectricCost) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadHeavyequipment()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\4005.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve HeavyequipmentMatrix(i)
			Input(10, HeavyequipmentMatrix(i).ID)
			Input(10, HeavyequipmentMatrix(i).TL)
			Input(10, HeavyequipmentMatrix(i).Weight)
			Input(10, HeavyequipmentMatrix(i).Volume)
			Input(10, HeavyequipmentMatrix(i).Cost)
			Input(10, HeavyequipmentMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadMedical()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\4006.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve MedicalMatrix(i)
			Input(10, MedicalMatrix(i).ID)
			Input(10, MedicalMatrix(i).TL)
			Input(10, MedicalMatrix(i).Weight)
			Input(10, MedicalMatrix(i).Volume)
			Input(10, MedicalMatrix(i).Cost)
			Input(10, MedicalMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadEntertainment()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\4007.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve EntertainmentMatrix(i)
			Input(10, EntertainmentMatrix(i).ID)
			Input(10, EntertainmentMatrix(i).TL)
			Input(10, EntertainmentMatrix(i).Weight)
			Input(10, EntertainmentMatrix(i).Volume)
			Input(10, EntertainmentMatrix(i).Cost)
			Input(10, EntertainmentMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadVehicleAccess()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\4008.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve VehicleAccessMatrix(i)
			Input(10, VehicleAccessMatrix(i).ID)
			Input(10, VehicleAccessMatrix(i).TL)
			Input(10, VehicleAccessMatrix(i).Weight)
			Input(10, VehicleAccessMatrix(i).Volume)
			Input(10, VehicleAccessMatrix(i).Cost)
			Input(10, VehicleAccessMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadTeleporters()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\4009.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve TeleporterMatrix(i)
			Input(10, TeleporterMatrix(i).ID)
			Input(10, TeleporterMatrix(i).TL)
			Input(10, TeleporterMatrix(i).Weight)
			Input(10, TeleporterMatrix(i).Volume)
			Input(10, TeleporterMatrix(i).Cost)
			Input(10, TeleporterMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadSecurity()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\4010.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve SecurityMatrix(i)
			Input(10, SecurityMatrix(i).ID)
			Input(10, SecurityMatrix(i).TL)
			Input(10, SecurityMatrix(i).Weight)
			Input(10, SecurityMatrix(i).Volume)
			Input(10, SecurityMatrix(i).Cost)
			Input(10, SecurityMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadVehicleStorage()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\4011.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve VehicleStorageMatrix(i)
			Input(10, VehicleStorageMatrix(i).ID)
			Input(10, VehicleStorageMatrix(i).Weight)
			Input(10, VehicleStorageMatrix(i).Volume)
			Input(10, VehicleStorageMatrix(i).Cost) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadLandingAids()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\4012.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve LandingAidMatrix(i)
			Input(10, LandingAidMatrix(i).ID)
			Input(10, LandingAidMatrix(i).TL)
			Input(10, LandingAidMatrix(i).Weight)
			Input(10, LandingAidMatrix(i).Volume)
			Input(10, LandingAidMatrix(i).Cost)
			Input(10, LandingAidMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadFuelAccessories()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\4013.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve FuelAccessoryMatrix(i)
			Input(10, FuelAccessoryMatrix(i).ID)
			Input(10, FuelAccessoryMatrix(i).TL)
			Input(10, FuelAccessoryMatrix(i).Weight)
			Input(10, FuelAccessoryMatrix(i).Volume)
			Input(10, FuelAccessoryMatrix(i).Cost)
			Input(10, FuelAccessoryMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadNuclearDampers()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\4014.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve NuclearDamperMatrix(i)
			Input(10, NuclearDamperMatrix(i).ID)
			Input(10, NuclearDamperMatrix(i).TL)
			Input(10, NuclearDamperMatrix(i).Weight)
			Input(10, NuclearDamperMatrix(i).Volume)
			Input(10, NuclearDamperMatrix(i).Cost)
			Input(10, NuclearDamperMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadRealityStabilizers()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\4015.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve RealityStabilizerMatrix(i)
			Input(10, RealityStabilizerMatrix(i).ID)
			Input(10, RealityStabilizerMatrix(i).TL)
			Input(10, RealityStabilizerMatrix(i).Weight)
			Input(10, RealityStabilizerMatrix(i).Volume)
			Input(10, RealityStabilizerMatrix(i).Cost)
			Input(10, RealityStabilizerMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadModularSockets()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\4016.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve ModularSocketMatrix(i)
			Input(10, ModularSocketMatrix(i).ID)
			Input(10, ModularSocketMatrix(i).TL)
			Input(10, ModularSocketMatrix(i).Weight)
			Input(10, ModularSocketMatrix(i).Volume)
			Input(10, ModularSocketMatrix(i).Cost)
			Input(10, ModularSocketMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
End Module