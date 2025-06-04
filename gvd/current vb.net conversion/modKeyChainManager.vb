Option Strict Off
Option Explicit On
Module modKeyChainManager 'make sure all keychain arrays start with a base of 1 and not 0
	
	Public Function VariantArrayToStringArray(ByVal KeyChain As Object) As String()
		Dim temparray() As String
		Dim i As Integer
		
		If IsArray(KeyChain) Then
			
			'UPGRADE_WARNING: Couldn't resolve default property of object KeyChain(1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If KeyChain(1) = "" Then
				'UPGRADE_WARNING: Lower bound of array temparray was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
				ReDim temparray(1)
				VariantArrayToStringArray = VB6.CopyArray(temparray)
				Exit Function
			Else
				For i = 1 To UBound(KeyChain)
					'UPGRADE_WARNING: Lower bound of array temparray was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
					ReDim Preserve temparray(i)
					'UPGRADE_WARNING: Couldn't resolve default property of object KeyChain(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					temparray(i) = KeyChain(i)
				Next 
			End If
		Else
			'UPGRADE_WARNING: Lower bound of array temparray was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim temparray(1)
			VariantArrayToStringArray = VB6.CopyArray(temparray)
		End If
		VariantArrayToStringArray = VB6.CopyArray(temparray)
		
	End Function
	
	Public Function mAddKey(ByVal KeyChain As Object, ByVal PropulsionKey As String) As String()
		'adds the key of a propulsion system to the END of the Keychain
		Dim NewSize As Integer
		
		If IsArray(KeyChain) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object KeyChain(1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If KeyChain(1) = "" Then
				NewSize = 1
			Else
				NewSize = UBound(KeyChain) + 1
			End If
		Else
			'UPGRADE_WARNING: Lower bound of array KeyChain was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
			ReDim KeyChain(1)
			'UPGRADE_WARNING: Couldn't resolve default property of object KeyChain(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			KeyChain(1) = PropulsionKey
			mAddKey = VariantArrayToStringArray(KeyChain)
			Exit Function
		End If
		
		'UPGRADE_WARNING: Lower bound of array KeyChain was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
		ReDim Preserve KeyChain(NewSize)
		'UPGRADE_WARNING: Couldn't resolve default property of object KeyChain(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		KeyChain(NewSize) = PropulsionKey
		
		mAddKey = VariantArrayToStringArray(KeyChain)
	End Function
	
	Public Function mRemoveKey(ByVal KeyChain As Object, ByVal PropulsionKey As String) As String()
		'removes a specified key from the keychain
		'then redimensions the keychain so that there are no empty positions in the chain
		'its important that there are no empty spaces (with the exception of the first key)
		'in the chain since the statistics calculations
		'that loop through the keychain expect a valid key at each location
		Dim chainsize As Integer
		Dim i As Integer
		Dim deletedkey As Integer
		
		'if the keychain is empty, then pass the empty keychain and exit the function
		'UPGRADE_WARNING: Couldn't resolve default property of object KeyChain(1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If KeyChain(1) = "" Then
			mRemoveKey = VariantArrayToStringArray(KeyChain)
			Exit Function
		Else
			chainsize = UBound(KeyChain) 'store the size of our keychain
			'loop through the keychain and find the key we need to delete
			For i = 1 To chainsize
				'UPGRADE_WARNING: Couldn't resolve default property of object KeyChain(i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If KeyChain(i) = PropulsionKey Then
					'UPGRADE_WARNING: Couldn't resolve default property of object KeyChain(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					KeyChain(i) = "" 'delete the key by setting it equal to ""
					deletedkey = i
				End If
			Next 
			If deletedkey > 0 Then 'MPJ 07/25/2000 the key does not exist for some reason! Probably due to bug in earlier version of GVD and the key never got added to keychain in first place
				If deletedkey = chainsize Then 'if we've reached the last position on the keychain we can redimension it
					If chainsize = 1 Then
						'minimum keychain size must be 1
						'UPGRADE_WARNING: Lower bound of array KeyChain was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
						ReDim KeyChain(1)
					Else
						'UPGRADE_WARNING: Lower bound of array KeyChain was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
						ReDim Preserve KeyChain(chainsize - 1)
					End If
				Else 'we need to move each key past the deleted key's location up one position
					For i = deletedkey To chainsize - 1
						'UPGRADE_WARNING: Couldn't resolve default property of object KeyChain(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						KeyChain(i) = KeyChain(i + 1)
					Next 
					'UPGRADE_WARNING: Lower bound of array KeyChain was changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="0F1C9BE1-AF9D-476E-83B1-17D43BECFF20"'
					ReDim Preserve KeyChain(chainsize - 1)
				End If
			End If
		End If
		
		mRemoveKey = VariantArrayToStringArray(KeyChain)
	End Function
	
	Public Sub mRemoveReferencedKeys(ByVal KeyChain As Object, ByVal Key As String, ByVal MethodName As String)
		Dim m_oCurrentVeh As Object
		'tell each of the power plants that
		'referenced this power using component to remove this key from its
		'keychain
		Dim chainsize As Integer
		Dim i As Integer
		
		chainsize = UBound(KeyChain)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object KeyChain(1). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (chainsize = 1) And (KeyChain(1) = "") Then
		Else
			For i = 1 To chainsize
				'todo: obsolete, since i believe i no longer use keychains.  Just remember to remove the
				' VEHICLE_DLL conditional compilation value in the vehicle.dll project properties / Make tab after
				' deleting this entire sub routine
#If VEHICLE_DLL Then
				'UPGRADE_NOTE: #If #EndIf block was not upgraded because the expression VEHICLE_DLL did not evaluate to True or was not evaluated. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="27EE2C3C-05AF-4C04-B2AF-657B4FB6B5FC"'
				CallByName Veh.Components(KeyChain(i)), MethodName, VbMethod, Key
#Else
				'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				CallByName(m_oCurrentVeh.Components(KeyChain(i)), MethodName, CallType.Method, Key)
#End If
			Next 
		End If
	End Sub
End Module