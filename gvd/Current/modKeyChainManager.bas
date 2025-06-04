Attribute VB_Name = "modKeyChainManager"
Option Explicit
Option Base 1 'make sure all keychain arrays start with a base of 1 and not 0

Public Function VariantArrayToStringArray(ByVal KeyChain As Variant) As String()
    Dim temparray() As String
    Dim i As Long
    
    If IsArray(KeyChain) Then
    
        If KeyChain(1) = "" Then
            ReDim temparray(1)
            VariantArrayToStringArray = temparray
            Exit Function
        Else
            For i = 1 To UBound(KeyChain)
                ReDim Preserve temparray(i)
                temparray(i) = KeyChain(i)
            Next
        End If
    Else
        ReDim temparray(1)
        VariantArrayToStringArray = temparray
    End If
    VariantArrayToStringArray = temparray

End Function

Public Function mAddKey(ByVal KeyChain As Variant, ByVal PropulsionKey As String) As String()
    'adds the key of a propulsion system to the END of the Keychain
    Dim NewSize As Long
    
    If IsArray(KeyChain) Then
        If KeyChain(1) = "" Then
            NewSize = 1
        Else
            NewSize = UBound(KeyChain) + 1
        End If
    Else
        ReDim KeyChain(1)
        KeyChain(1) = PropulsionKey
        mAddKey = VariantArrayToStringArray(KeyChain)
        Exit Function
    End If
    
     ReDim Preserve KeyChain(1 To NewSize)
     KeyChain(NewSize) = PropulsionKey
    
    mAddKey = VariantArrayToStringArray(KeyChain)
End Function

Public Function mRemoveKey(ByVal KeyChain As Variant, ByVal PropulsionKey As String) As String()
    'removes a specified key from the keychain
    'then redimensions the keychain so that there are no empty positions in the chain
    'its important that there are no empty spaces (with the exception of the first key)
    'in the chain since the statistics calculations
    'that loop through the keychain expect a valid key at each location
    Dim chainsize As Long
    Dim i As Long
    Dim deletedkey As Long
    
    'if the keychain is empty, then pass the empty keychain and exit the function
    If KeyChain(1) = "" Then
        mRemoveKey = VariantArrayToStringArray(KeyChain)
        Exit Function
    Else
        chainsize = UBound(KeyChain) 'store the size of our keychain
        'loop through the keychain and find the key we need to delete
        For i = 1 To chainsize
            If KeyChain(i) = PropulsionKey Then
                KeyChain(i) = "" 'delete the key by setting it equal to ""
                deletedkey = i
            End If
        Next
        If deletedkey > 0 Then 'MPJ 07/25/2000 the key does not exist for some reason! Probably due to bug in earlier version of GVD and the key never got added to keychain in first place
            If deletedkey = chainsize Then 'if we've reached the last position on the keychain we can redimension it
                If chainsize = 1 Then
                    'minimum keychain size must be 1
                    ReDim KeyChain(1)
                Else
                    ReDim Preserve KeyChain(chainsize - 1)
                End If
            Else 'we need to move each key past the deleted key's location up one position
                For i = deletedkey To chainsize - 1
                    KeyChain(i) = KeyChain(i + 1)
                Next
                ReDim Preserve KeyChain(chainsize - 1)
            End If
        End If
    End If
    
    mRemoveKey = VariantArrayToStringArray(KeyChain)
End Function

Public Sub mRemoveReferencedKeys(ByVal KeyChain As Variant, ByVal Key As String, ByVal MethodName As String)
    'tell each of the power plants that
    'referenced this power using component to remove this key from its
    'keychain
    Dim chainsize As Long
    Dim i As Long
    
    chainsize = UBound(KeyChain)
    
    If (chainsize = 1) And (KeyChain(1) = "") Then
    Else
        For i = 1 To chainsize
        'todo: obsolete, since i believe i no longer use keychains.  Just remember to remove the
        ' VEHICLE_DLL conditional compilation value in the vehicle.dll project properties / Make tab after
        ' deleting this entire sub routine
            #If VEHICLE_DLL Then
                CallByName Veh.Components(KeyChain(i)), MethodName, VbMethod, Key
            #Else
                CallByName m_oCurrentVeh.Components(KeyChain(i)), MethodName, VbMethod, Key
            #End If
        Next
    End If
End Sub
