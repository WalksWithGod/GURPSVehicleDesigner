Attribute VB_Name = "modTableMiscEquipt"
Option Explicit

Public Type udtArmMotor
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type
    
Public Type udtBilgePump
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type
    
Public Type udtFireExtinguisher
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
End Type
    
Public Type udtWorkshop
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
    PhysicsMod As Single
    ElectricMod As Single
    ElectricMod2 As Single
    ElectricCost As Single
End Type

Public Type udtHeavyequipment
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtMedical
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtEntertainment
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtVehicleAccess
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtTeleporter
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtSecurity
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtVehicleStorage
    ID As Integer
    Weight As Single
    Volume As Single
    Cost As Single
End Type

Public Type udtLandingAid
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtFuelAccessory
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtNuclearDamper
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtRealityStabilizer
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtModularSocket
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type


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
Dim i As Integer
Open App.Path & "\data\4001.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve ArmMotorMatrix(i)
            Input #10, ArmMotorMatrix(i).ID, ArmMotorMatrix(i).TL, ArmMotorMatrix(i).Weight, ArmMotorMatrix(i).Volume, ArmMotorMatrix(i).Cost, ArmMotorMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadBilgePumps()
Dim i As Integer
Open App.Path & "\data\4002.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve BilgePumpMatrix(i)
            Input #10, BilgePumpMatrix(i).ID, BilgePumpMatrix(i).TL, BilgePumpMatrix(i).Weight, BilgePumpMatrix(i).Volume, BilgePumpMatrix(i).Cost, BilgePumpMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadFireExtinguishers()
Dim i As Integer
Open App.Path & "\data\4003.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve FireExtinguisherMatrix(i)
            Input #10, FireExtinguisherMatrix(i).ID, FireExtinguisherMatrix(i).TL, FireExtinguisherMatrix(i).Weight, FireExtinguisherMatrix(i).Volume, FireExtinguisherMatrix(i).Cost ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadWorkshops()
Dim i As Integer
Open App.Path & "\data\4004.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve WorkshopMatrix(i)
            Input #10, WorkshopMatrix(i).ID, WorkshopMatrix(i).TL, WorkshopMatrix(i).Weight, WorkshopMatrix(i).Volume, WorkshopMatrix(i).Cost, WorkshopMatrix(i).Power, WorkshopMatrix(i).PhysicsMod, WorkshopMatrix(i).ElectricMod, WorkshopMatrix(i).ElectricMod2, WorkshopMatrix(i).ElectricCost ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadHeavyequipment()
Dim i As Integer
Open App.Path & "\data\4005.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve HeavyequipmentMatrix(i)
            Input #10, HeavyequipmentMatrix(i).ID, HeavyequipmentMatrix(i).TL, HeavyequipmentMatrix(i).Weight, HeavyequipmentMatrix(i).Volume, HeavyequipmentMatrix(i).Cost, HeavyequipmentMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadMedical()
Dim i As Integer
Open App.Path & "\data\4006.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve MedicalMatrix(i)
            Input #10, MedicalMatrix(i).ID, MedicalMatrix(i).TL, MedicalMatrix(i).Weight, MedicalMatrix(i).Volume, MedicalMatrix(i).Cost, MedicalMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadEntertainment()
Dim i As Integer
Open App.Path & "\data\4007.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve EntertainmentMatrix(i)
            Input #10, EntertainmentMatrix(i).ID, EntertainmentMatrix(i).TL, EntertainmentMatrix(i).Weight, EntertainmentMatrix(i).Volume, EntertainmentMatrix(i).Cost, EntertainmentMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadVehicleAccess()
Dim i As Integer
Open App.Path & "\data\4008.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve VehicleAccessMatrix(i)
            Input #10, VehicleAccessMatrix(i).ID, VehicleAccessMatrix(i).TL, VehicleAccessMatrix(i).Weight, VehicleAccessMatrix(i).Volume, VehicleAccessMatrix(i).Cost, VehicleAccessMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadTeleporters()
Dim i As Integer
Open App.Path & "\data\4009.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve TeleporterMatrix(i)
            Input #10, TeleporterMatrix(i).ID, TeleporterMatrix(i).TL, TeleporterMatrix(i).Weight, TeleporterMatrix(i).Volume, TeleporterMatrix(i).Cost, TeleporterMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadSecurity()
Dim i As Integer
Open App.Path & "\data\4010.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve SecurityMatrix(i)
            Input #10, SecurityMatrix(i).ID, SecurityMatrix(i).TL, SecurityMatrix(i).Weight, SecurityMatrix(i).Volume, SecurityMatrix(i).Cost, SecurityMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadVehicleStorage()
Dim i As Integer
Open App.Path & "\data\4011.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve VehicleStorageMatrix(i)
            Input #10, VehicleStorageMatrix(i).ID, VehicleStorageMatrix(i).Weight, VehicleStorageMatrix(i).Volume, VehicleStorageMatrix(i).Cost  ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadLandingAids()
Dim i As Integer
Open App.Path & "\data\4012.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve LandingAidMatrix(i)
            Input #10, LandingAidMatrix(i).ID, LandingAidMatrix(i).TL, LandingAidMatrix(i).Weight, LandingAidMatrix(i).Volume, LandingAidMatrix(i).Cost, LandingAidMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadFuelAccessories()
Dim i As Integer
Open App.Path & "\data\4013.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve FuelAccessoryMatrix(i)
            Input #10, FuelAccessoryMatrix(i).ID, FuelAccessoryMatrix(i).TL, FuelAccessoryMatrix(i).Weight, FuelAccessoryMatrix(i).Volume, FuelAccessoryMatrix(i).Cost, FuelAccessoryMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadNuclearDampers()
Dim i As Integer
Open App.Path & "\data\4014.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve NuclearDamperMatrix(i)
            Input #10, NuclearDamperMatrix(i).ID, NuclearDamperMatrix(i).TL, NuclearDamperMatrix(i).Weight, NuclearDamperMatrix(i).Volume, NuclearDamperMatrix(i).Cost, NuclearDamperMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadRealityStabilizers()
Dim i As Integer
Open App.Path & "\data\4015.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve RealityStabilizerMatrix(i)
            Input #10, RealityStabilizerMatrix(i).ID, RealityStabilizerMatrix(i).TL, RealityStabilizerMatrix(i).Weight, RealityStabilizerMatrix(i).Volume, RealityStabilizerMatrix(i).Cost, RealityStabilizerMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadModularSockets()
Dim i As Integer
Open App.Path & "\data\4016.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve ModularSocketMatrix(i)
            Input #10, ModularSocketMatrix(i).ID, ModularSocketMatrix(i).TL, ModularSocketMatrix(i).Weight, ModularSocketMatrix(i).Volume, ModularSocketMatrix(i).Cost, ModularSocketMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub
