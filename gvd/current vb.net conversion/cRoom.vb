Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("cRoom_NET.cRoom")> Public Class cRoom
	
	' cRoom.cls
	' A room is defined as _ANY_ enclosed internal space.
	' However, the difference between a room and a compartment and a container?
	' is that a container represents the "hull" of the vehicle and
	' has external walls and internal structure (e.g. bulk heads, beams, etc)
	
	' Practically speaking, a turret, pod, superstructure are all containers
	' whereas cargo, cabins, suites, galleys, holoventure zones are all rooms.
	' Hell, these could even be convertible to the different types of rooms if
	' a remodel job was done.
	
	' Incidentally, there is no difference between these different types of rooms
	' as far as the simulation is concerned except that some rooms have
	' sub-components in them which provide
	' bonuses such as +5 sleep recovery for sleeping in a cabin.  Sleeping
	' would be allowed in a galley, but it wouldnt provide a bonus.  Galley
	' provides +10 to food production however (these values are all made up and
	' for demonstration purposes only).  A cabin though might also have a small
	' food bonus depending on the cabin (maybe it has a fridge or small kitchen
	' or a replicator).
	
	'Implements cIComponent
	'Implements cIDisplay
	'Implements cINode
	'Implements cIPersist
	'Implements cIBUild
End Class