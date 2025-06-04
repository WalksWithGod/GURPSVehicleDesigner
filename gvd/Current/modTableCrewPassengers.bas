Attribute VB_Name = "modTableCrewPassengers"
Option Explicit

Public Type udtManeuverControl
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
End Type

Public Type udtCrewStation
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
End Type

Public Type udtAccommodations
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
End Type

Public Type udtLimitedLifeSystem
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtFullLifeSystem
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
    WeightPP As Single
    VolumePP As Single
    CostPP As Single
    PowerPP As Single
End Type

Public Type udtArtificialGravUnit
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtSafety
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
    Power As Single
End Type

Public Type udtProvisions
    ID As Integer
    TL As Integer
    Weight As Single
    Volume As Single
    Cost As Single
End Type

Public ManeuverControlMatrix() As udtManeuverControl
Public CrewStationMatrix() As udtCrewStation
Public AccommodationsMatrix() As udtAccommodations
Public LimitedLifeSystemMatrix() As udtLimitedLifeSystem
Public FullLifeSystemMatrix() As udtFullLifeSystem
Public ArtificialGravUnitMatrix() As udtArtificialGravUnit
Public SafetyMatrix() As udtSafety
Public ProvisionsMatrix() As udtProvisions


Sub LoadManeuverControls()
Dim i As Integer
Open App.Path & "\data\5001.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve ManeuverControlMatrix(i)
            Input #10, ManeuverControlMatrix(i).ID, ManeuverControlMatrix(i).TL, ManeuverControlMatrix(i).Weight, ManeuverControlMatrix(i).Volume, ManeuverControlMatrix(i).Cost ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadCrewStations()
Dim i As Integer
Open App.Path & "\data\5002.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve CrewStationMatrix(i)
            Input #10, CrewStationMatrix(i).ID, CrewStationMatrix(i).TL, CrewStationMatrix(i).Weight, CrewStationMatrix(i).Volume, CrewStationMatrix(i).Cost ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadAccommodations()
Dim i As Integer
Open App.Path & "\data\5003.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve AccommodationsMatrix(i)
            Input #10, AccommodationsMatrix(i).ID, AccommodationsMatrix(i).TL, AccommodationsMatrix(i).Weight, AccommodationsMatrix(i).Volume, AccommodationsMatrix(i).Cost ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub



Sub LoadLimitedLifeSystems()
Dim i As Integer
Open App.Path & "\data\5004.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve LimitedLifeSystemMatrix(i)
            Input #10, LimitedLifeSystemMatrix(i).ID, LimitedLifeSystemMatrix(i).TL, LimitedLifeSystemMatrix(i).Weight, LimitedLifeSystemMatrix(i).Volume, LimitedLifeSystemMatrix(i).Cost, LimitedLifeSystemMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadFullLifeSystems()
Dim i As Integer
Open App.Path & "\data\5005.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve FullLifeSystemMatrix(i)
            Input #10, FullLifeSystemMatrix(i).ID, FullLifeSystemMatrix(i).TL, FullLifeSystemMatrix(i).Weight, FullLifeSystemMatrix(i).Volume, FullLifeSystemMatrix(i).Cost, FullLifeSystemMatrix(i).Power, FullLifeSystemMatrix(i).WeightPP, FullLifeSystemMatrix(i).VolumePP, FullLifeSystemMatrix(i).CostPP, FullLifeSystemMatrix(i).PowerPP ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadArtificialGravUnits()
Dim i As Integer
Open App.Path & "\data\5006.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve ArtificialGravUnitMatrix(i)
            Input #10, ArtificialGravUnitMatrix(i).ID, ArtificialGravUnitMatrix(i).TL, ArtificialGravUnitMatrix(i).Weight, ArtificialGravUnitMatrix(i).Volume, ArtificialGravUnitMatrix(i).Cost, ArtificialGravUnitMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadSafety()
Dim i As Integer
Open App.Path & "\data\5007.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve SafetyMatrix(i)
            Input #10, SafetyMatrix(i).ID, SafetyMatrix(i).TL, SafetyMatrix(i).Weight, SafetyMatrix(i).Volume, SafetyMatrix(i).Cost, SafetyMatrix(i).Power ' Read data into the udt
        i = i + 1
Loop
Close #10    ' Close file.
End Sub

Sub LoadProvisions()
Dim i As Integer
Open App.Path & "\data\5008.txt" For Input As #10 ' Open file for input.
i = 1 ' intialize the counter
Do While Not EOF(10) ' Loop until end of file.
        ReDim Preserve ProvisionsMatrix(i)
            Input #10, ProvisionsMatrix(i).ID, ProvisionsMatrix(i).TL, ProvisionsMatrix(i).Weight, ProvisionsMatrix(i).Volume, ProvisionsMatrix(i).Cost
        i = i + 1
Loop
Close #10    ' Close file.
End Sub


