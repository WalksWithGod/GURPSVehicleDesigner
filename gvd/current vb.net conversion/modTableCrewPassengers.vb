Option Strict Off
Option Explicit On
Module modTableCrewPassengers
	
	Public Structure udtManeuverControl
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
	End Structure
	
	Public Structure udtCrewStation
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
	End Structure
	
	Public Structure udtAccommodations
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
	End Structure
	
	Public Structure udtLimitedLifeSystem
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtFullLifeSystem
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
		Dim WeightPP As Single
		Dim VolumePP As Single
		Dim CostPP As Single
		Dim PowerPP As Single
	End Structure
	
	Public Structure udtArtificialGravUnit
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtSafety
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
		Dim Power As Single
	End Structure
	
	Public Structure udtProvisions
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
	End Structure
	
	Public ManeuverControlMatrix() As udtManeuverControl
	Public CrewStationMatrix() As udtCrewStation
	Public AccommodationsMatrix() As udtAccommodations
	Public LimitedLifeSystemMatrix() As udtLimitedLifeSystem
	Public FullLifeSystemMatrix() As udtFullLifeSystem
	Public ArtificialGravUnitMatrix() As udtArtificialGravUnit
	Public SafetyMatrix() As udtSafety
	Public ProvisionsMatrix() As udtProvisions
	
	
	Sub LoadManeuverControls()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\5001.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve ManeuverControlMatrix(i)
			Input(10, ManeuverControlMatrix(i).ID)
			Input(10, ManeuverControlMatrix(i).TL)
			Input(10, ManeuverControlMatrix(i).Weight)
			Input(10, ManeuverControlMatrix(i).Volume)
			Input(10, ManeuverControlMatrix(i).Cost) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadCrewStations()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\5002.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve CrewStationMatrix(i)
			Input(10, CrewStationMatrix(i).ID)
			Input(10, CrewStationMatrix(i).TL)
			Input(10, CrewStationMatrix(i).Weight)
			Input(10, CrewStationMatrix(i).Volume)
			Input(10, CrewStationMatrix(i).Cost) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadAccommodations()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\5003.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve AccommodationsMatrix(i)
			Input(10, AccommodationsMatrix(i).ID)
			Input(10, AccommodationsMatrix(i).TL)
			Input(10, AccommodationsMatrix(i).Weight)
			Input(10, AccommodationsMatrix(i).Volume)
			Input(10, AccommodationsMatrix(i).Cost) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	
	
	Sub LoadLimitedLifeSystems()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\5004.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve LimitedLifeSystemMatrix(i)
			Input(10, LimitedLifeSystemMatrix(i).ID)
			Input(10, LimitedLifeSystemMatrix(i).TL)
			Input(10, LimitedLifeSystemMatrix(i).Weight)
			Input(10, LimitedLifeSystemMatrix(i).Volume)
			Input(10, LimitedLifeSystemMatrix(i).Cost)
			Input(10, LimitedLifeSystemMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadFullLifeSystems()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\5005.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve FullLifeSystemMatrix(i)
			Input(10, FullLifeSystemMatrix(i).ID)
			Input(10, FullLifeSystemMatrix(i).TL)
			Input(10, FullLifeSystemMatrix(i).Weight)
			Input(10, FullLifeSystemMatrix(i).Volume)
			Input(10, FullLifeSystemMatrix(i).Cost)
			Input(10, FullLifeSystemMatrix(i).Power)
			Input(10, FullLifeSystemMatrix(i).WeightPP)
			Input(10, FullLifeSystemMatrix(i).VolumePP)
			Input(10, FullLifeSystemMatrix(i).CostPP)
			Input(10, FullLifeSystemMatrix(i).PowerPP) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadArtificialGravUnits()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\5006.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve ArtificialGravUnitMatrix(i)
			Input(10, ArtificialGravUnitMatrix(i).ID)
			Input(10, ArtificialGravUnitMatrix(i).TL)
			Input(10, ArtificialGravUnitMatrix(i).Weight)
			Input(10, ArtificialGravUnitMatrix(i).Volume)
			Input(10, ArtificialGravUnitMatrix(i).Cost)
			Input(10, ArtificialGravUnitMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadSafety()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\5007.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve SafetyMatrix(i)
			Input(10, SafetyMatrix(i).ID)
			Input(10, SafetyMatrix(i).TL)
			Input(10, SafetyMatrix(i).Weight)
			Input(10, SafetyMatrix(i).Volume)
			Input(10, SafetyMatrix(i).Cost)
			Input(10, SafetyMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadProvisions()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\5008.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve ProvisionsMatrix(i)
			Input(10, ProvisionsMatrix(i).ID)
			Input(10, ProvisionsMatrix(i).TL)
			Input(10, ProvisionsMatrix(i).Weight)
			Input(10, ProvisionsMatrix(i).Volume)
			Input(10, ProvisionsMatrix(i).Cost)
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
End Module