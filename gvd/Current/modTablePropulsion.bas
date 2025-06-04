Attribute VB_Name = "modTablePropulsion"
Option Explicit

Public Type udtHarness
    ID As Integer
    TL As Integer
    Weight As Single
    FlyingMultiplier As Single
    SwimmingMultiplier As Single
    Cost As Single
    Volume As Single
    Efficiency As Single
    Power As Single
End Type

Public Type udtSail
    ID As Integer
    TL As Integer
    WeightCloth As Single
    WeightSynthetic As Single
    WeightBioplas As Single
    MotiveThrust As Single
    CostCloth As Single
    CostSynthetic As Single
    CostBioplas As Single
    Volume As Single
    Power As Single
End Type

Public Type udtRowing
    ID As Integer
    TL As Integer
    Weight As Single
    RowingMultiplier As Single
    Cost As Single
    Volume As Single
    Power As Single
End Type

Public Type udtlightSail
    ID As Integer
    TL As Integer
    Weight As Single
    ThrustSquareModifier As Single
    Cost As Single
    Volume As Single
    Thrust As Single
    Power As Single
End Type

Public Type udtDrivetrain
    ID As Integer
    TL As Integer
    Weight1 As Single
    Weight2 As Single
    Weight3 As Single
    Cost As Single
    Volume As Single
End Type

Public Type udtAquatic
    ID As Integer
    TL As Integer
    Weight1 As Single
    Weight2 As Single
    Weight3 As Single
    MotiveThrust As Single
    Cost As Single
    Volume As Single
End Type

Public Type udtAirscrew
    ID As Integer
    TL As Integer
    Weight1 As Single
    Weight2 As Single
    Weight3 As Single
    MotiveThrust As Single
    Cost As Single
    Volume As Single
End Type

Public Type udtHelicopter
    ID As Integer
    TL As Integer
    Lift As Single
    MotiveThrust As Single
    Weight1 As Single
    Weight2 As Single
    Weight3 As Single
    Cost As Single
    Volume As Single
End Type

Public Type udtOrnithopter
    ID As Integer
    TL As Integer
    MotiveThrust As Single
    Lift As Single
    Weight1 As Single
    Weight2 As Single
    Weight3 As Single
    Cost As Single
    Volume As Single
End Type

Public Type udtJet
    ID As Integer
    TL As Integer
    Weight1 As Single
    Weight2 As Single
    Cost As Single
    Volume As Single
    ABWeight As Single
    ABCost As Single
    ABThrust As Single
    ABFuel As Single
    Fuel As Single
    Power As Single
End Type

Public Type udtRocket
    ID As Integer
    TL As Integer
    Weight1 As Single
    Weight2 As Single
    Cost As Single
    Volume As Single
    Fuel As Single
    Power As Single
End Type

Public Type udtOrionEngine
    ID As Integer
    TL As Integer
    Weight1 As Single
    Weight2 As Single
    Weight3 As Single
    MotiveThrust As Single
    Volume As Single
    Power As Single
End Type

Public Type udtThrustBomb
    ID As Integer
    TL As Integer
    Weight1 As Single
    Weight2 As Single
    Cost As Single
    Volume As Single
    Power As Single
End Type

Public Type udtSolidRocket
    ID As Integer
    TL As Integer
    Weight As Single
    Cost As Single
    Volume As Single
    Power As Single
End Type

Public Type udtReactionless
    ID As Integer
    TL As Integer
    Weight As Single
    Cost As Single
    MinCost As Single
    Volume As Single
    Power As Single
End Type

Public Type udtMagLev
    ID As Integer
    TL As Integer
    Weight As Single
    Cost As Single
    Volume As Single
    Power As Single
End Type

Public Type udtStarDrve
    ID As Integer
    TL As Integer
    Weight1 As Single
    Weight2 As Single
    Cost1 As Single
    Cost2 As Single
    Volume As Single
    Power1 As Single
    Power2 As Single
End Type

Public Type udtLiftingGas
    ID As Integer
    TL As Integer
    Cost As Single
    Volume As Single
    Power As Single
End Type

Public Type udtContraGravity
    ID As Integer
    TL As Integer
    Weight1 As Single
    Weight2 As Single
    Cost1 As Single
    Cost2 As Single
    Volume As Single
    Power As Single
End Type
    
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
Dim i As Integer
Open App.Path & "\data\1001.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve HarnessMatrix(i)
            Input #10, HarnessMatrix(i).ID, HarnessMatrix(i).TL, HarnessMatrix(i).Weight, HarnessMatrix(i).FlyingMultiplier, HarnessMatrix(i).SwimmingMultiplier, HarnessMatrix(i).Cost, HarnessMatrix(i).Volume, HarnessMatrix(i).Efficiency, HarnessMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadSail()
Dim i As Integer
Open App.Path & "\data\1002.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve SailMatrix(i)
            Input #10, SailMatrix(i).ID, SailMatrix(i).TL, SailMatrix(i).WeightCloth, SailMatrix(i).WeightSynthetic, SailMatrix(i).WeightBioplas, SailMatrix(i).MotiveThrust, SailMatrix(i).CostCloth, SailMatrix(i).CostSynthetic, SailMatrix(i).CostBioplas, SailMatrix(i).Volume, SailMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub
    
Sub LoadRowing()
Dim i As Integer
Open App.Path & "\data\1003.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve RowingMatrix(i)
            Input #10, RowingMatrix(i).ID, RowingMatrix(i).TL, RowingMatrix(i).Weight, RowingMatrix(i).RowingMultiplier, RowingMatrix(i).Cost, RowingMatrix(i).Volume, RowingMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadlightSail()
Dim i As Integer
Open App.Path & "\data\1004.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve lightSailMatrix(i)
            Input #10, lightSailMatrix(i).ID, lightSailMatrix(i).TL, lightSailMatrix(i).Weight, lightSailMatrix(i).ThrustSquareModifier, lightSailMatrix(i).Cost, lightSailMatrix(i).Volume, lightSailMatrix(i).Thrust, lightSailMatrix(i).Power  ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub


Sub LoadDrivetrain()
Dim i As Integer
Open App.Path & "\data\1005.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve DrivetrainMatrix(i)
            Input #10, DrivetrainMatrix(i).ID, DrivetrainMatrix(i).TL, DrivetrainMatrix(i).Weight1, DrivetrainMatrix(i).Weight2, DrivetrainMatrix(i).Weight3, DrivetrainMatrix(i).Cost, DrivetrainMatrix(i).Volume ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadAquatic()
Dim i As Integer
Open App.Path & "\data\1006.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve AquaticMatrix(i)
            Input #10, AquaticMatrix(i).ID, AquaticMatrix(i).TL, AquaticMatrix(i).Weight1, AquaticMatrix(i).Weight2, AquaticMatrix(i).Weight3, AquaticMatrix(i).MotiveThrust, AquaticMatrix(i).Cost, AquaticMatrix(i).Volume ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadAirscrew()
Dim i As Integer
Open App.Path & "\data\1007.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve AirscrewMatrix(i)
            Input #10, AirscrewMatrix(i).ID, AirscrewMatrix(i).TL, AirscrewMatrix(i).Weight1, AirscrewMatrix(i).Weight2, AirscrewMatrix(i).Weight3, AirscrewMatrix(i).MotiveThrust, AirscrewMatrix(i).Cost, AirscrewMatrix(i).Volume ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadHelicopter()
Dim i As Integer
Open App.Path & "\data\1008.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve HelicopterMatrix(i)
            Input #10, HelicopterMatrix(i).ID, HelicopterMatrix(i).TL, HelicopterMatrix(i).Lift, HelicopterMatrix(i).MotiveThrust, HelicopterMatrix(i).Weight1, HelicopterMatrix(i).Weight2, HelicopterMatrix(i).Weight3, HelicopterMatrix(i).Cost, HelicopterMatrix(i).Volume ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadOrnithopter()
Dim i As Integer
Open App.Path & "\data\1009.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve OrnithopterMatrix(i)
            Input #10, OrnithopterMatrix(i).ID, OrnithopterMatrix(i).TL, OrnithopterMatrix(i).MotiveThrust, OrnithopterMatrix(i).Lift, OrnithopterMatrix(i).Weight1, OrnithopterMatrix(i).Weight2, OrnithopterMatrix(i).Weight3, OrnithopterMatrix(i).Cost, OrnithopterMatrix(i).Volume ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadJet()
Dim i As Integer
Open App.Path & "\data\1010.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve JetMatrix(i)
            Input #10, JetMatrix(i).ID, JetMatrix(i).TL, JetMatrix(i).Weight1, JetMatrix(i).Weight2, JetMatrix(i).Cost, JetMatrix(i).Volume, JetMatrix(i).ABWeight, JetMatrix(i).ABCost, JetMatrix(i).ABThrust, JetMatrix(i).ABFuel, JetMatrix(i).Fuel, JetMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadRocket()
Dim i As Integer
Open App.Path & "\data\1011.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve RocketMatrix(i)
            Input #10, RocketMatrix(i).ID, RocketMatrix(i).TL, RocketMatrix(i).Weight1, RocketMatrix(i).Weight2, RocketMatrix(i).Cost, RocketMatrix(i).Volume, RocketMatrix(i).Fuel, RocketMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub


Sub LoadOrionEngine()
Dim i As Integer
Open App.Path & "\data\1012.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve OrionEngineMatrix(i)
            Input #10, OrionEngineMatrix(i).ID, OrionEngineMatrix(i).TL, OrionEngineMatrix(i).Weight1, OrionEngineMatrix(i).Weight2, OrionEngineMatrix(i).Weight3, OrionEngineMatrix(i).MotiveThrust, OrionEngineMatrix(i).Volume, OrionEngineMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadThrustBomb()
Dim i As Integer
Open App.Path & "\data\1013.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve ThrustBombMatrix(i)
            Input #10, ThrustBombMatrix(i).ID, ThrustBombMatrix(i).TL, ThrustBombMatrix(i).Weight1, ThrustBombMatrix(i).Weight2, ThrustBombMatrix(i).Cost, ThrustBombMatrix(i).Volume, ThrustBombMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadSolidRocket()
Dim i As Integer
Open App.Path & "\data\1014.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve SolidRocketMatrix(i)
            Input #10, SolidRocketMatrix(i).ID, SolidRocketMatrix(i).TL, SolidRocketMatrix(i).Weight, SolidRocketMatrix(i).Cost, SolidRocketMatrix(i).Volume, SolidRocketMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadReactionless()
Dim i As Integer
Open App.Path & "\data\1015.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve ReactionlessMatrix(i)
            Input #10, ReactionlessMatrix(i).ID, ReactionlessMatrix(i).TL, ReactionlessMatrix(i).Weight, ReactionlessMatrix(i).Cost, ReactionlessMatrix(i).MinCost, ReactionlessMatrix(i).Volume, ReactionlessMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadMagLev()
Dim i As Integer
Open App.Path & "\data\1016.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve MagLevMatrix(i)
            Input #10, MagLevMatrix(i).ID, MagLevMatrix(i).TL, MagLevMatrix(i).Weight, MagLevMatrix(i).Cost, MagLevMatrix(i).Volume, MagLevMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadStarDrive()
Dim i As Integer
Open App.Path & "\data\1017.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve StarDriveMatrix(i)
            Input #10, StarDriveMatrix(i).ID, StarDriveMatrix(i).TL, StarDriveMatrix(i).Weight1, StarDriveMatrix(i).Weight2, StarDriveMatrix(i).Cost1, StarDriveMatrix(i).Cost2, StarDriveMatrix(i).Volume, StarDriveMatrix(i).Power1, StarDriveMatrix(i).Power2  ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadLiftingGas()
Dim i As Integer
Open App.Path & "\data\1018.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve LiftingGasMatrix(i)
            Input #10, LiftingGasMatrix(i).ID, LiftingGasMatrix(i).TL, LiftingGasMatrix(i).Cost, LiftingGasMatrix(i).Volume, LiftingGasMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadContraGravity()
Dim i As Integer
Open App.Path & "\data\1019.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve ContraGravMatrix(i)
            Input #10, ContraGravMatrix(i).ID, ContraGravMatrix(i).TL, ContraGravMatrix(i).Weight1, ContraGravMatrix(i).Weight2, ContraGravMatrix(i).Cost1, ContraGravMatrix(i).Cost2, ContraGravMatrix(i).Volume, ContraGravMatrix(i).Power  ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub
