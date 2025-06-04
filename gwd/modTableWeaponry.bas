Attribute VB_Name = "modTableWeaponry"
Option Explicit

Public Type udtHardPoint
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
End Type
    

Public Type udtWeaponAccessory
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
End Type
    
Public Type udtAmmo
    Name As String
    Damage1 As String
    Damage2 As String
    Fragmentation As Boolean
    Formula As String
    Multiplier As Single
    Divisor As Single
    Range As Single
    WPS As Single
    CPS As Single
    Accuracy As Single
End Type

Public Type udtGuidanceSystem
    Name As String
    Brilliant As Boolean
    TL As Integer
    WeightMod As Single
    CostMod As Single
    Skill As String
End Type

Public GuidanceMatrix() As udtGuidanceSystem
Public AmmoMatrix() As udtAmmo
Public HardpointMatrix() As udtHardPoint
Public WeaponAccessoryMatrix() As udtWeaponAccessory

Sub LoadGuidanceSystems()
Dim i As Integer
Open App.Path & "\data\7003.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the co unter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve GuidanceMatrix(i)
            Input #10, GuidanceMatrix(i).Name, GuidanceMatrix(i).Brilliant, GuidanceMatrix(i).TL, GuidanceMatrix(i).WeightMod, GuidanceMatrix(i).CostMod, GuidanceMatrix(i).Skill   ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub



Sub LoadAmmunitionTable()
Dim i As Integer
Open App.Path & "\data\Ammunition.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve AmmoMatrix(i)
            Input #10, AmmoMatrix(i).Name, AmmoMatrix(i).Damage1, AmmoMatrix(i).Damage2, AmmoMatrix(i).Fragmentation, AmmoMatrix(i).Formula, AmmoMatrix(i).Multiplier, AmmoMatrix(i).Divisor, AmmoMatrix(i).Range, AmmoMatrix(i).WPS, AmmoMatrix(i).CPS, AmmoMatrix(i).Accuracy  ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadWeaponAccessories()
Dim i As Integer
Open App.Path & "\data\7001.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve WeaponAccessoryMatrix(i)
            Input #10, WeaponAccessoryMatrix(i).ID, WeaponAccessoryMatrix(i).TL, WeaponAccessoryMatrix(i).Weight, WeaponAccessoryMatrix(i).Volume, WeaponAccessoryMatrix(i).Cost ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub



Sub LoadHardPoints()
Dim i As Integer
Open App.Path & "\data\7002.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve HardpointMatrix(i)
            Input #10, HardpointMatrix(i).ID, HardpointMatrix(i).TL, HardpointMatrix(i).Weight, HardpointMatrix(i).Volume, HardpointMatrix(i).Cost ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

