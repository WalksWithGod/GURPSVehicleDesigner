Attribute VB_Name = "modTableLoader"
Option Explicit
Public Type udtArmor
    MaterialType As Integer
    TL As Integer
    Quality As Integer
    WeightMod As Single
    Cost As Single
End Type

Public Type udtSurfaceArea
    Volume As Single
    Area As Single
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

Public Type udtStoneBoltDamage
    Str As Long
    Thrust1 As Long
    Thrust2 As Long
    Swing1 As Long
    Swing2 As Long
End Type

Public SurfaceAreaMatrix() As udtSurfaceArea
Public StoneBoltDamageMatrix() As udtStoneBoltDamage
Public SurfaceMatrix() As udtSurfaceFeature
Public ArmorMatrix() As udtArmor
Public GroundStabMatrix() As udtGroundStability
Public WaterStabMatrix(12, 1 To 6) As udtWaterStability


Sub LoadSurfaceAreaTable()

Dim i As Integer
Open App.Path & "\data\surfaceareatable.txt" For Input As #4 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(4) ' Loop until end of file.
        ReDim Preserve SurfaceAreaMatrix(i)
        Input #4, SurfaceAreaMatrix(i).Volume, SurfaceAreaMatrix(i).Area
        i = i + 1
Loop
Close #4    ' Close file.
End Sub

Sub LoadSurfaceFeatureTable()

Dim i As Integer
Open App.Path & "\data\surfacetable.txt" For Input As #4 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(4) ' Loop until end of file.
        ReDim Preserve SurfaceMatrix(i)
        Input #4, SurfaceMatrix(i).TL, SurfaceMatrix(i).FeatureType, SurfaceMatrix(i).WeightMod, SurfaceMatrix(i).CostMod, SurfaceMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #4    ' Close file.
End Sub

Sub LoadArmorTable()
Dim i As Integer
Dim k As Integer 'counter 2
Dim j As Integer ' counter 3
Open App.Path & "\data\armortable.txt" For Input As #3 ' Open file for input.
i = 1 ' intialize the counter
k = 2 ' initalize counter 2
j = 1 ' init counter 3

Do While Not EOF(3) ' Loop until end of file.
        ReDim Preserve ArmorMatrix(i + 9)
        Input #3, ArmorMatrix(i).MaterialType, ArmorMatrix(i).Quality, ArmorMatrix(i).WeightMod, ArmorMatrix(i + 1).WeightMod, ArmorMatrix(i + 2).WeightMod, ArmorMatrix(i + 3).WeightMod, ArmorMatrix(i + 4).WeightMod, ArmorMatrix(i + 5).WeightMod, ArmorMatrix(i + 6).WeightMod, ArmorMatrix(i + 7).WeightMod, ArmorMatrix(i + 8).WeightMod, ArmorMatrix(i + 9).WeightMod, ArmorMatrix(i).Cost ' Read data into the udt
            ArmorMatrix(i).TL = 4
            For k = k To k + 8

                'If k = 242 Then Exit Sub 'exit before reaching end of file
                
                ArmorMatrix(k).MaterialType = ArmorMatrix(i).MaterialType
                ArmorMatrix(k).Quality = ArmorMatrix(i).Quality
                ArmorMatrix(k).TL = ArmorMatrix(i).TL + j
                ArmorMatrix(k).Cost = ArmorMatrix(i).Cost
                j = j + 1
                If j = 10 Then j = 1
            Next
            i = i + 10
            k = k + 1
            
Loop
Close #3    ' Close file.
End Sub
Sub LoadGroundStabilityTable()

Dim i As Integer
Open App.Path & "\data\gMRandgSR.txt" For Input As #5 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(5) ' Loop until end of file.
        ReDim Preserve GroundStabMatrix(i)
        Input #5, GroundStabMatrix(i).MotiveSystem, GroundStabMatrix(i).M1, GroundStabMatrix(i).S1, GroundStabMatrix(i).M2, GroundStabMatrix(i).S2, GroundStabMatrix(i).M3, GroundStabMatrix(i).S3, GroundStabMatrix(i).M4, GroundStabMatrix(i).S4, GroundStabMatrix(i).M5, GroundStabMatrix(i).S5 ' Read data into the udt
        i = i + 1
Loop
Close #5    ' Close file.
End Sub

Sub LoadWaterStabilityTable()

Dim i As Integer
Open App.Path & "\data\wMRandwSR.txt" For Input As #6 ' Open file for input.
i = 0 ' intialize the counter
Do While Not EOF(6) ' Loop until end of file.
        Input #6, WaterStabMatrix(i, 1).MR, WaterStabMatrix(i, 1).SR, WaterStabMatrix(i, 2).MR, WaterStabMatrix(i, 2).SR, WaterStabMatrix(i, 3).MR, WaterStabMatrix(i, 3).SR, WaterStabMatrix(i, 4).MR, WaterStabMatrix(i, 4).SR, WaterStabMatrix(i, 5).MR, WaterStabMatrix(i, 5).SR, WaterStabMatrix(i, 6).MR, WaterStabMatrix(i, 6).SR ' Read data into the udt
        i = i + 1
Loop
Close #6    ' Close file.
End Sub

Sub LoadStoneBoltDamageTable()

Dim i As Integer
Open App.Path & "\data\StoneBoltDamage.txt" For Input As #6 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(6) ' Loop until end of file.
        ReDim Preserve StoneBoltDamageMatrix(i)
        Input #6, StoneBoltDamageMatrix(i).Str, StoneBoltDamageMatrix(i).Thrust1, StoneBoltDamageMatrix(i).Thrust2, StoneBoltDamageMatrix(i).Swing1, StoneBoltDamageMatrix(i).Swing2
        i = i + 1
Loop
Close #6    ' Close file.
End Sub

