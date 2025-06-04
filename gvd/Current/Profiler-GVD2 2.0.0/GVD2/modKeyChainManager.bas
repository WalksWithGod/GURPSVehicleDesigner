Attribute VB_Name = "modKeyChainManager"
Option Explicit
Option Base 1 'make sure all keychain arrays start with a base of 1 and not 0

Public Function VariantArrayToStringArray(ByVal KeyChain As Variant) As String()
vbwProfiler.vbwProcIn 634
    Dim temparray() As String
    Dim i As Long

vbwProfiler.vbwExecuteLine 11442
    If IsArray(KeyChain) Then

vbwProfiler.vbwExecuteLine 11443
        If KeyChain(1) = "" Then
vbwProfiler.vbwExecuteLine 11444
            ReDim temparray(1)
vbwProfiler.vbwExecuteLine 11445
            VariantArrayToStringArray = temparray
vbwProfiler.vbwProcOut 634
vbwProfiler.vbwExecuteLine 11446
            Exit Function
        Else
vbwProfiler.vbwExecuteLine 11447 'B
vbwProfiler.vbwExecuteLine 11448
            For i = 1 To UBound(KeyChain)
vbwProfiler.vbwExecuteLine 11449
                ReDim Preserve temparray(i)
vbwProfiler.vbwExecuteLine 11450
                temparray(i) = KeyChain(i)
vbwProfiler.vbwExecuteLine 11451
            Next
        End If
vbwProfiler.vbwExecuteLine 11452 'B
    Else
vbwProfiler.vbwExecuteLine 11453 'B
vbwProfiler.vbwExecuteLine 11454
        ReDim temparray(1)
vbwProfiler.vbwExecuteLine 11455
        VariantArrayToStringArray = temparray
    End If
vbwProfiler.vbwExecuteLine 11456 'B
vbwProfiler.vbwExecuteLine 11457
    VariantArrayToStringArray = temparray

vbwProfiler.vbwProcOut 634
vbwProfiler.vbwExecuteLine 11458
End Function

Public Function mAddKey(ByVal KeyChain As Variant, ByVal PropulsionKey As String) As String()
    'adds the key of a propulsion system to the END of the Keychain
vbwProfiler.vbwProcIn 635
    Dim NewSize As Long

vbwProfiler.vbwExecuteLine 11459
    If IsArray(KeyChain) Then
vbwProfiler.vbwExecuteLine 11460
        If KeyChain(1) = "" Then
vbwProfiler.vbwExecuteLine 11461
            NewSize = 1
        Else
vbwProfiler.vbwExecuteLine 11462 'B
vbwProfiler.vbwExecuteLine 11463
            NewSize = UBound(KeyChain) + 1
        End If
vbwProfiler.vbwExecuteLine 11464 'B
    Else
vbwProfiler.vbwExecuteLine 11465 'B
vbwProfiler.vbwExecuteLine 11466
        ReDim KeyChain(1)
vbwProfiler.vbwExecuteLine 11467
        KeyChain(1) = PropulsionKey
vbwProfiler.vbwExecuteLine 11468
        mAddKey = VariantArrayToStringArray(KeyChain)
vbwProfiler.vbwProcOut 635
vbwProfiler.vbwExecuteLine 11469
        Exit Function
    End If
vbwProfiler.vbwExecuteLine 11470 'B

vbwProfiler.vbwExecuteLine 11471
     ReDim Preserve KeyChain(1 To NewSize)
vbwProfiler.vbwExecuteLine 11472
     KeyChain(NewSize) = PropulsionKey

vbwProfiler.vbwExecuteLine 11473
    mAddKey = VariantArrayToStringArray(KeyChain)
vbwProfiler.vbwProcOut 635
vbwProfiler.vbwExecuteLine 11474
End Function

Public Function mRemoveKey(ByVal KeyChain As Variant, ByVal PropulsionKey As String) As String()
    'removes a specified key from the keychain
    'then redimensions the keychain so that there are no empty positions in the chain
    'its important that there are no empty spaces (with the exception of the first key)
    'in the chain since the statistics calculations
    'that loop through the keychain expect a valid key at each location
vbwProfiler.vbwProcIn 636
    Dim chainsize As Long
    Dim i As Long
    Dim deletedkey As Long

    'if the keychain is empty, then pass the empty keychain and exit the function
vbwProfiler.vbwExecuteLine 11475
    If KeyChain(1) = "" Then
vbwProfiler.vbwExecuteLine 11476
        mRemoveKey = VariantArrayToStringArray(KeyChain)
vbwProfiler.vbwProcOut 636
vbwProfiler.vbwExecuteLine 11477
        Exit Function
    Else
vbwProfiler.vbwExecuteLine 11478 'B
vbwProfiler.vbwExecuteLine 11479
        chainsize = UBound(KeyChain) 'store the size of our keychain
        'loop through the keychain and find the key we need to delete
vbwProfiler.vbwExecuteLine 11480
        For i = 1 To chainsize
vbwProfiler.vbwExecuteLine 11481
            If KeyChain(i) = PropulsionKey Then
vbwProfiler.vbwExecuteLine 11482
                KeyChain(i) = "" 'delete the key by setting it equal to ""
vbwProfiler.vbwExecuteLine 11483
                deletedkey = i
            End If
vbwProfiler.vbwExecuteLine 11484 'B
vbwProfiler.vbwExecuteLine 11485
        Next
vbwProfiler.vbwExecuteLine 11486
        If deletedkey > 0 Then 'MPJ 07/25/2000 the key does not exist for some reason! Probably due to bug in earlier version of GVD and the key never got added to keychain in first place
vbwProfiler.vbwExecuteLine 11487
            If deletedkey = chainsize Then 'if we've reached the last position on the keychain we can redimension it
vbwProfiler.vbwExecuteLine 11488
                If chainsize = 1 Then
                    'minimum keychain size must be 1
vbwProfiler.vbwExecuteLine 11489
                    ReDim KeyChain(1)
                Else
vbwProfiler.vbwExecuteLine 11490 'B
vbwProfiler.vbwExecuteLine 11491
                    ReDim Preserve KeyChain(chainsize - 1)
                End If
vbwProfiler.vbwExecuteLine 11492 'B
            Else 'we need to move each key past the deleted key's location up one position
vbwProfiler.vbwExecuteLine 11493 'B
vbwProfiler.vbwExecuteLine 11494
                For i = deletedkey To chainsize - 1
vbwProfiler.vbwExecuteLine 11495
                    KeyChain(i) = KeyChain(i + 1)
vbwProfiler.vbwExecuteLine 11496
                Next
vbwProfiler.vbwExecuteLine 11497
                ReDim Preserve KeyChain(chainsize - 1)
            End If
vbwProfiler.vbwExecuteLine 11498 'B
        End If
vbwProfiler.vbwExecuteLine 11499 'B
    End If
vbwProfiler.vbwExecuteLine 11500 'B

vbwProfiler.vbwExecuteLine 11501
    mRemoveKey = VariantArrayToStringArray(KeyChain)
vbwProfiler.vbwProcOut 636
vbwProfiler.vbwExecuteLine 11502
End Function

Public Sub mRemoveReferencedKeys(ByVal KeyChain As Variant, ByVal Key As String, ByVal MethodName As String)
    'tell each of the power plants that
    'referenced this power using component to remove this key from its
    'keychain
vbwProfiler.vbwProcIn 637
    Dim chainsize As Long
    Dim i As Long

vbwProfiler.vbwExecuteLine 11503
    chainsize = UBound(KeyChain)

vbwProfiler.vbwExecuteLine 11504
    If (chainsize = 1) And (KeyChain(1) = "") Then
    Else
vbwProfiler.vbwExecuteLine 11505 'B
vbwProfiler.vbwExecuteLine 11506
        For i = 1 To chainsize
        'todo: obsolete, since i believe i no longer use keychains.  Just remember to remove the
        ' VEHICLE_DLL conditional compilation value in the vehicle.dll project properties / Make tab after
        ' deleting this entire sub routine
            #If VEHICLE_DLL Then
                CallByName Veh.Components(KeyChain(i)), MethodName, VbMethod, Key
            #Else
vbwProfiler.vbwExecuteLine 11507
                CallByName m_oCurrentVeh.Components(KeyChain(i)), MethodName, VbMethod, Key
            #End If
vbwProfiler.vbwExecuteLine 11508
        Next
    End If
vbwProfiler.vbwExecuteLine 11509 'B
vbwProfiler.vbwProcOut 637
vbwProfiler.vbwExecuteLine 11510
End Sub


