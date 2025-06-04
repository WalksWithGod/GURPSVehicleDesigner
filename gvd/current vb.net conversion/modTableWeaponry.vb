Option Strict Off
Option Explicit On
Module modTableWeaponry
	
	Public Structure udtHardPoint
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
	End Structure
	
	
	Public Structure udtWeaponAccessory
		Dim ID As Short
		Dim TL As Short
		Dim Weight As Single
		Dim Volume As Single
		Dim Cost As Single
	End Structure
	
	Public Structure udtAmmo
		Dim Name As String
		Dim Damage1 As String
		Dim Damage2 As String
		Dim Fragmentation As Boolean
		Dim Formula As String
		Dim Multiplier As Single
		Dim Divisor As Single
		Dim Range As Single
		Dim WPS As Single
		Dim CPS As Single
		Dim Accuracy As Single
	End Structure
	
	Public Structure udtGuidanceSystem
		Dim Name As String
		Dim Brilliant As Boolean
		Dim TL As Short
		Dim WeightMod As Single
		Dim CostMod As Single
		Dim Skill As String
	End Structure
	
	Public GuidanceMatrix() As udtGuidanceSystem
	Public AmmoMatrix() As udtAmmo
	Public HardpointMatrix() As udtHardPoint
	Public WeaponAccessoryMatrix() As udtWeaponAccessory
	
	Sub LoadGuidanceSystems()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\7003.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the co unter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve GuidanceMatrix(i)
			Input(10, GuidanceMatrix(i).Name)
			Input(10, GuidanceMatrix(i).Brilliant)
			Input(10, GuidanceMatrix(i).TL)
			Input(10, GuidanceMatrix(i).WeightMod)
			Input(10, GuidanceMatrix(i).CostMod)
			Input(10, GuidanceMatrix(i).Skill) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	
	
	Sub LoadAmmunitionTable()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\Ammunition.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve AmmoMatrix(i)
			Input(10, AmmoMatrix(i).Name)
			Input(10, AmmoMatrix(i).Damage1)
			Input(10, AmmoMatrix(i).Damage2)
			Input(10, AmmoMatrix(i).Fragmentation)
			Input(10, AmmoMatrix(i).Formula)
			Input(10, AmmoMatrix(i).Multiplier)
			Input(10, AmmoMatrix(i).Divisor)
			Input(10, AmmoMatrix(i).Range)
			Input(10, AmmoMatrix(i).WPS)
			Input(10, AmmoMatrix(i).CPS)
			Input(10, AmmoMatrix(i).Accuracy) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	Sub LoadWeaponAccessories()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\7001.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve WeaponAccessoryMatrix(i)
			Input(10, WeaponAccessoryMatrix(i).ID)
			Input(10, WeaponAccessoryMatrix(i).TL)
			Input(10, WeaponAccessoryMatrix(i).Weight)
			Input(10, WeaponAccessoryMatrix(i).Volume)
			Input(10, WeaponAccessoryMatrix(i).Cost) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
	
	
	
	Sub LoadHardPoints()
		Dim i As Short
		FileOpen(10, My.Application.Info.DirectoryPath & "\data\7002.txt", OpenMode.Input) ' Open file for input.
		i = 1 ' intialize the counter
		Do While Not EOF(10) ' Loop until end of file.
			ReDim Preserve HardpointMatrix(i)
			Input(10, HardpointMatrix(i).ID)
			Input(10, HardpointMatrix(i).TL)
			Input(10, HardpointMatrix(i).Weight)
			Input(10, HardpointMatrix(i).Volume)
			Input(10, HardpointMatrix(i).Cost) ' Read data into the udt
			i = i + 1
		Loop 
		FileClose(10) ' Close file.
	End Sub
End Module