Attribute VB_Name = "modTableLoader"
Option Explicit
Public Type udtArmor
    MaterialType As Integer
    TL As Integer
    Quality As Integer
    WeightMod As Single
    Cost As Single
End Type

Public Type udtSurfaceFeature
    FeatureType As Integer
    TL As Integer
    WeightMod As Single
    CostMod As Single
    Power As Integer
End Type

Public Type udtGroundStability
    MotiveSystem As Integer
    M1 As Single
    S1 As Integer
    M2 As Single
    S2 As Integer
    M3 As Single
    S3 As Integer
    M4 As Single
    S4 As Integer
    M5 As Single
    S5 As Integer
End Type

Public Type udtWaterStability
    MR As Single
    SR As Integer
End Type

Public SurfaceMatrix(1 To 56) As udtSurfaceFeature
Public ArmorMatrix(1 To 240) As udtArmor
Public GroundStabMatrix(1 To 13) As udtGroundStability
Public WaterStabMatrix(12, 1 To 6) As udtWaterStability

Sub LoadSurfaceFeatureTable()

Dim i As Integer 'counter 1
Open "surfacetable.txt" For Input As #4 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(4) ' Loop until end of file.
        Input #4, SurfaceMatrix(i).TL, SurfaceMatrix(i).FeatureType, SurfaceMatrix(i).WeightMod, SurfaceMatrix(i).CostMod, SurfaceMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #4    ' Close file.
End Sub

Sub LoadArmorTable()
Dim i As Integer 'counter 1
Dim k As Integer 'counter 2
Open "armortable.txt" For Input As #3 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(3) ' Loop until end of file.
        Input #3, ArmorMatrix(i).MaterialType, ArmorMatrix(i).Quality, ArmorMatrix(i).WeightMod, ArmorMatrix(i + 1).WeightMod, ArmorMatrix(i + 2).WeightMod, ArmorMatrix(i + 3).WeightMod, ArmorMatrix(i + 4).WeightMod, ArmorMatrix(i + 5).WeightMod, ArmorMatrix(i + 6).WeightMod, ArmorMatrix(i + 7).WeightMod, ArmorMatrix(i + 8).WeightMod, ArmorMatrix(i + 9).WeightMod, ArmorMatrix(i).Cost ' Read data into the udt
            ArmorMatrix(i).TL = 4
            For k = 2 To 10
            ArmorMatrix(k).MaterialType = ArmorMatrix(i).MaterialType
            ArmorMatrix(k).Quality = ArmorMatrix(i).Quality
            ArmorMatrix(k).TL = ArmorMatrix(i).TL + k - 1
            ArmorMatrix(k).Cost = ArmorMatrix(i).Cost
            Next
            i = i + 10
Loop
Close #3    ' Close file.
End Sub

Sub LoadGroundStabilityTable()

Dim i As Integer 'counter 1
Open "gMRandgSR.txt" For Input As #5 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(5) ' Loop until end of file.
        Input #5, GroundStabMatrix(i).MotiveSystem, GroundStabMatrix(i).M1, GroundStabMatrix(i).S1, GroundStabMatrix(i).M2, GroundStabMatrix(i).S2, GroundStabMatrix(i).M3, GroundStabMatrix(i).S3, GroundStabMatrix(i).M4, GroundStabMatrix(i).S4, GroundStabMatrix(i).M5, GroundStabMatrix(i).S5 ' Read data into the udt
        i = i + 1
Loop
Close #5    ' Close file.
End Sub

Sub LoadWaterStabilityTable()

Dim i As Integer 'counter 1
Open "wMRandwSR.txt" For Input As #6 ' Open file for input.
i = 0 ' intialize the counter
Do While Not EOF(6) ' Loop until end of file.
        Input #6, WaterStabMatrix(i, 1).MR, WaterStabMatrix(i, 1).SR, WaterStabMatrix(i, 2).MR, WaterStabMatrix(i, 2).SR, WaterStabMatrix(i, 3).MR, WaterStabMatrix(i, 3).SR, WaterStabMatrix(i, 4).MR, WaterStabMatrix(i, 4).SR, WaterStabMatrix(i, 5).MR, WaterStabMatrix(i, 5).SR, WaterStabMatrix(i, 6).MR, WaterStabMatrix(i, 6).SR ' Read data into the udt
        i = i + 1
Loop
Close #6    ' Close file.
End Sub
