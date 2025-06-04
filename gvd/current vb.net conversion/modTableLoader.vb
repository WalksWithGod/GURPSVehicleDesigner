Option Strict Off
Option Explicit On
Module modTableLoader
	Public Structure udtArmor
		Dim MaterialType As Short
		Dim TL As Short
		Dim Quality As Short
		Dim WeightMod As Single
		Dim Cost As Single
	End Structure
	
	Public Structure udtSurfaceArea
		Dim Volume As Single
		Dim Area As Single
	End Structure
	
	Public Structure udtSurfaceFeature
		Dim FeatureType As Short
		Dim TL As Short
		Dim WeightMod As Single
		Dim CostMod As Single
		Dim Power As Short
	End Structure
	
	Public Structure udtGroundStability
		Dim MotiveSystem As Short
		Dim M1 As Single
		Dim S1 As Short
		Dim M2 As Single
		Dim S2 As Short
		Dim M3 As Single
		Dim S3 As Short
		Dim M4 As Single
		Dim S4 As Short
		Dim M5 As Single
		Dim S5 As Short
	End Structure
	
	Public Structure udtWaterStability
		Dim MR As Single
		Dim SR As Short
	End Structure
	
	Public Structure udtStoneBoltDamage
		'UPGRADE_NOTE: Str was upgraded to Str_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Str_Renamed As Integer
		Dim Thrust1 As Integer
		Dim Thrust2 As Integer
		Dim Swing1 As Integer
		Dim Swing2 As Integer
	End Structure
	
	Public SurfaceAreaMatrix() As udtSurfaceArea
	Public StoneBoltDamageMatrix() As udtStoneBoltDamage
	Public SurfaceMatrix() As udtSurfaceFeature
	Public ArmorMatrix() As udtArmor
	Public GroundStabMatrix() As udtGroundStability
	'UPGRADE_WARNING: Lower bound of array WaterStabMatrix was changed from 0,1 to 0,0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
	Public WaterStabMatrix(12, 6) As udtWaterStability
	
	
	Sub LoadSurfaceAreaTable()
		
		Dim i As Short
		FileOpen(4, My.Application.Info.DirectoryPath & "\data\surfaceareatable.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(4) ' Loop until end of file.
			ReDim Preserve SurfaceAreaMatrix(i)
			Input(4, SurfaceAreaMatrix(i).Volume)
			Input(4, SurfaceAreaMatrix(i).Area)
			i = i + 1
		Loop 
		FileClose(4) ' Close file.
	End Sub
	
	Sub LoadSurfaceFeatureTable()
		
		Dim i As Short
		FileOpen(4, My.Application.Info.DirectoryPath & "\data\surfacetable.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(4) ' Loop until end of file.
			ReDim Preserve SurfaceMatrix(i)
			Input(4, SurfaceMatrix(i).TL)
			Input(4, SurfaceMatrix(i).FeatureType)
			Input(4, SurfaceMatrix(i).WeightMod)
			Input(4, SurfaceMatrix(i).CostMod)
			Input(4, SurfaceMatrix(i).Power) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(4) ' Close file.
	End Sub
	
	Sub LoadArmorTable()
		Dim i As Short
		Dim k As Short 'counter 2
		Dim j As Short ' counter 3
		FileOpen(3, My.Application.Info.DirectoryPath & "\data\armortable.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		k = 2 ' initalize counter 2
		j = 1 ' init counter 3
		
		Do While Not EOF(3) ' Loop until end of file.
			ReDim Preserve ArmorMatrix(i + 9)
			Input(3, ArmorMatrix(i).MaterialType)
			Input(3, ArmorMatrix(i).Quality)
			Input(3, ArmorMatrix(i).WeightMod)
			Input(3, ArmorMatrix(i + 1).WeightMod)
			Input(3, ArmorMatrix(i + 2).WeightMod)
			Input(3, ArmorMatrix(i + 3).WeightMod)
			Input(3, ArmorMatrix(i + 4).WeightMod)
			Input(3, ArmorMatrix(i + 5).WeightMod)
			Input(3, ArmorMatrix(i + 6).WeightMod)
			Input(3, ArmorMatrix(i + 7).WeightMod)
			Input(3, ArmorMatrix(i + 8).WeightMod)
			Input(3, ArmorMatrix(i + 9).WeightMod)
			Input(3, ArmorMatrix(i).Cost) ' Read data into the udt
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
		FileClose(3) ' Close file.
	End Sub
	Sub LoadGroundStabilityTable()
		
		Dim i As Short
		FileOpen(5, My.Application.Info.DirectoryPath & "\data\gMRandgSR.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(5) ' Loop until end of file.
			ReDim Preserve GroundStabMatrix(i)
			Input(5, GroundStabMatrix(i).MotiveSystem)
			Input(5, GroundStabMatrix(i).M1)
			Input(5, GroundStabMatrix(i).S1)
			Input(5, GroundStabMatrix(i).M2)
			Input(5, GroundStabMatrix(i).S2)
			Input(5, GroundStabMatrix(i).M3)
			Input(5, GroundStabMatrix(i).S3)
			Input(5, GroundStabMatrix(i).M4)
			Input(5, GroundStabMatrix(i).S4)
			Input(5, GroundStabMatrix(i).M5)
			Input(5, GroundStabMatrix(i).S5) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(5) ' Close file.
	End Sub
	
	Sub LoadWaterStabilityTable()
		
		Dim i As Short
		FileOpen(6, My.Application.Info.DirectoryPath & "\data\wMRandwSR.txt", OpenMode.Input) ' Open file for input.
		i = 0 ' intialize the counter
		Do While Not EOF(6) ' Loop until end of file.
			Input(6, WaterStabMatrix(i, 1).MR)
			Input(6, WaterStabMatrix(i, 1).SR)
			Input(6, WaterStabMatrix(i, 2).MR)
			Input(6, WaterStabMatrix(i, 2).SR)
			Input(6, WaterStabMatrix(i, 3).MR)
			Input(6, WaterStabMatrix(i, 3).SR)
			Input(6, WaterStabMatrix(i, 4).MR)
			Input(6, WaterStabMatrix(i, 4).SR)
			Input(6, WaterStabMatrix(i, 5).MR)
			Input(6, WaterStabMatrix(i, 5).SR)
			Input(6, WaterStabMatrix(i, 6).MR)
			Input(6, WaterStabMatrix(i, 6).SR) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(6) ' Close file.
	End Sub
	
	Sub LoadStoneBoltDamageTable()
		
		Dim i As Short
		FileOpen(6, My.Application.Info.DirectoryPath & "\data\StoneBoltDamage.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(6) ' Loop until end of file.
			ReDim Preserve StoneBoltDamageMatrix(i)
			Input(6, StoneBoltDamageMatrix(i).Str_Renamed)
			Input(6, StoneBoltDamageMatrix(i).Thrust1)
			Input(6, StoneBoltDamageMatrix(i).Thrust2)
			Input(6, StoneBoltDamageMatrix(i).Swing1)
			Input(6, StoneBoltDamageMatrix(i).Swing2)
			i = i + 1
		Loop 
		FileClose(6) ' Close file.
	End Sub
End Module