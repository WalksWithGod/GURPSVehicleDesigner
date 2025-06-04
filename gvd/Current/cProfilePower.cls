VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProfilePower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' This is a generic Profile class for handling EITHER Power Suppliers + Power Consumer associations or Fuel Suppliers and Fuel Consuming associations
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Private m_SC_Group() As clsSupplyConsumeGroup ' a single profile can and will usually have multiple SCGroups
Private m_sUnAssignedConsumerList() As String
Private sKey As String                        ' the name of this particular profile given by the user
Private m_TreeX As TreeX
Private m_lngGroupCount As Long
Private m_lngUnAssignedCount As Long
Private m_lngProfileType As Long
Private mvarDescription As String

Public Property Let UnAssignedConsumerList(ByRef s() As String)
    Dim sArr() As String ' this is so lame... Option Base 1 arrays when passed from a function turn dyanmic arrays into Fixed Arrays!
    Dim i As Long
    
    m_lngUnAssignedCount = UBound(s)
    ReDim m_sUnAssignedConsumerList(m_lngUnAssignedCount)
    
    ' this  hack keeps our m_sUnaassignedConsumerList dynamic and redim-able
    For i = 1 To m_lngUnAssignedCount
        m_sUnAssignedConsumerList(i) = s(i)
    Next
    
    ' cant use direct assignment cuz it turns our m_sUnassignedConsumerList into a fixed array.  This only happens when
    ' passing Option Base 1 arrays (even if they're dynamic!) to Subs/Functions and then trying to use them for assignment
    'm_sUnAssignedConsumerList = s
    
End Property

Public Function Group(ByVal iIndex As Long) As clsSupplyConsumeGroup
    Set Group = m_SC_Group(iIndex)
End Function

Public Property Get GroupCount() As Long
    GroupCount = m_lngGroupCount
End Property
Public Property Get UnAssignedConsumerList() As String()
    UnAssignedConsumerList = m_sUnAssignedConsumerList
End Property

Public Property Set Tree(ByRef O As Object)
    
    Set m_TreeX = O
End Property

Public Property Get ProfileType() As Long
    ProfileType = m_lngProfileType
End Property

Public Property Let ProfileType(ByVal l As Long)
    m_lngProfileType = l
End Property

Public Sub Show()
    Dim i As Long
    If Not m_TreeX Is Nothing Then
        m_TreeX.RemoveAllItems
        For i = 1 To m_lngGroupCount
            m_SC_Group(i).Show m_TreeX
        Next
    End If
    
    Veh.ActiveProfile = sKey
    Debug.Assert m_lngProfileType > 0
    Veh.ActiveProfileType = m_lngProfileType
End Sub

Friend Function SortAvailableSuppliers(ByRef vArr As Variant)
    ' actually this is stupid.  Every time a supplier or consumer is added to a vehicle,
    ' these should be added to every profile by calling
    ' Veh.AddConsumer or Veh.AddSupplier  <- the keymanager shouldnt be doing this since it doesnt
    ' likewise, when a supplier or consumer is removed, Veh.RemoveConsumer and Veh.RemoveSupplier should be called.
    ' HOWEVER, if a new profile is made, we automatically create a NEW group for each supplier.  But there is
    ' NO SORTING needed since each is initially put in its own group.
    ' need to manage these for anything else...
    ' However, we do need a Collection class to hold all profiles rather than using a standard collection
    ' Hrmm... but wait.  I do need to keep list of all suppliers and consumers in the keymanager also dont i?  Since
    ' when a new profile is created after alot of components have already been added, these all need to be
    ' received from the keymanager to the new profile.  Hrmm... think about this some more... trying to avoid redundant tracking of these
    ' stupid keys
    
    
    ' 05/15/02 MPJ - as stated above ("this is stupid") this function is not needed.  The way this should be handled
    ' is all suppliers and consumers are added to individual key arrays in the keymanager.  This is the only place these
    ' keys will be stored.
    ' this function receives all available suppliers in the vehicle and Creates new groups
    ' for those not already in a group
    
'    Dim i As Long
'    Dim s As String
'
'    If vArr(i) = "" Then Exit Function
'    For i = 1 To UBound(vArr)
'        For j = 1 To UBound(m_SC_Group.suppliers)
'    Next

End Function

Public Function AssignConsumer(ByRef s As String, ByVal lngGroupIndex As Long)
    Dim l As Long
    
    If RemoveUnAssignedConsumer(s) Then
        m_SC_Group(lngGroupIndex).AddConsumer (s)
    Else
        Debug.Print "clsProfilePower:AssignConsumer -- ERROR:  RemoveUnAssignedConsumer Returned FALSE"
    End If
End Function
Public Function MoveConsumer(lngConsumer, lngOldGroup, lngNewGroup)
    ' called to move a consumer from one SC_Group to another
    Dim sKey As String
    sKey = m_SC_Group(lngOldGroup).Consumer(lngConsumer)
    m_SC_Group(lngOldGroup).RemoveConsumerByIndex (lngConsumer)
    m_SC_Group(lngNewGroup).AddConsumer (sKey)
   
End Function

Public Function AddConsumer(ByRef s As String) As Long
    ' called when a new consuemr is added to the vehicle
    m_lngUnAssignedCount = UBound(m_sUnAssignedConsumerList)
    If (m_lngUnAssignedCount = 1) Then
        If (m_sUnAssignedConsumerList(1) = "") Then
            m_sUnAssignedConsumerList(m_lngUnAssignedCount) = s
            Exit Function
        End If
    End If
    
    m_lngUnAssignedCount = m_lngUnAssignedCount + 1
    
    ReDim Preserve m_sUnAssignedConsumerList(m_lngUnAssignedCount)
    m_sUnAssignedConsumerList(m_lngUnAssignedCount) = s
    Debug.Print "Consumer Added " & s
End Function

Public Function RemoveConsumer(ByRef s As String) As Long
    ' called when consumer is removed from entire vehicle
    Dim i As Long
    If Not RemoveUnAssignedConsumer(s) Then
        For i = 1 To m_lngGroupCount
            m_SC_Group(i).RemoveConsumer (s)
        Next
    End If
End Function

Public Function UnAssignConsumer(ByVal iGroupIndex As Long, ByVal iConsumerIndex As Long)
    Dim s As String
    
    s = m_SC_Group(iGroupIndex).Consumer(iConsumerIndex)
    m_SC_Group(iGroupIndex).RemoveConsumer (s)
    
    m_lngUnAssignedCount = m_lngUnAssignedCount + 1
    ReDim Preserve m_sUnAssignedConsumerList(m_lngUnAssignedCount)
    m_sUnAssignedConsumerList(m_lngUnAssignedCount) = s
    
End Function

Private Function RemoveUnAssignedConsumer(ByRef s As String) As Long
    ' returns TRUE if the element was found and deleted
    Dim i As Long
    Dim j As Long
    
    For i = 1 To m_lngUnAssignedCount - 1
        If s = m_sUnAssignedConsumerList(i) Then
            ' found it, now remove it
            For j = i To m_lngUnAssignedCount - 1
                m_sUnAssignedConsumerList(j) = m_sUnAssignedConsumerList(j + 1)
            Next
            RemoveUnAssignedConsumer = True
            m_lngUnAssignedCount = m_lngUnAssignedCount - 1
            ReDim Preserve m_sUnAssignedConsumerList(m_lngUnAssignedCount)
            Exit Function
        End If
    Next
    
    ' didnt find it, so maybe its the very last one in the array
    If s = m_sUnAssignedConsumerList(m_lngUnAssignedCount) Then
        m_lngUnAssignedCount = m_lngUnAssignedCount - 1
        ReDim Preserve m_sUnAssignedConsumerList(m_lngUnAssignedCount)
        RemoveUnAssignedConsumer = True
        Exit Function
    End If
    RemoveUnAssignedConsumer = False
    
End Function


Public Function MovePowerSystem(lngSupplier, lngOldGroup, lngNewGroup)
    ' called to move a power system from one SC_Group to another
    Dim sKey As String
    sKey = m_SC_Group(lngOldGroup).Supplier(lngSupplier)
    m_SC_Group(lngOldGroup).RemoveSupplierByIndex (lngSupplier)
    m_SC_Group(lngNewGroup).AddSupplier (sKey)
    Call RemoveEmptySCGroup
End Function

Public Function CreateNewSCGroup() As Long
    Dim O As clsSupplyConsumeGroup
    
    m_lngGroupCount = m_lngGroupCount + 1
    ReDim Preserve m_SC_Group(m_lngGroupCount)
    Set O = New clsSupplyConsumeGroup
    Set m_SC_Group(m_lngGroupCount) = O
    m_SC_Group(m_lngGroupCount).GroupIndex = m_lngGroupCount
    Set O = Nothing
    CreateNewSCGroup = m_lngGroupCount
End Function
Public Function AddPowerSystem(ByRef s As String) As Long
    ' called when a new power system has been added to the vehicle.  These power systems
    ' get added to a new SC_Group to which they are the only member
    Dim iNewGroup As Long
    
    iNewGroup = CreateNewSCGroup
    
    m_SC_Group(iNewGroup).AddSupplier s
    
End Function

Public Function RemovePowerSystem(ByRef s As String) As Long
    ' called when a power system is removed from entire vehicle
    Dim i As Long
    For i = 1 To UBound(m_SC_Group)
        m_SC_Group(i).RemoveSupplier s
    Next
    Call RemoveEmptySCGroup
End Function

Private Sub RemoveEmptySCGroup()
    Dim i As Long
    ' Only groups with BOTH no suppliers AND no consumers are deleted automatically for the user
    ' remove any empty sc groups IN REVERSE ORDER so that we dont have a situtation where i = 1 in first itteration then i = 2
    ' is pointing to 3rd element since first itteration caused a shift in the entire array
    For i = UBound(m_SC_Group) To 1 Step -1
        If (m_SC_Group(i).SupplierCount = 0) And (m_SC_Group(i).ConsumerCount = 0) Then
            RemoveSCGroup (i)
        End If
    Next
End Sub
Private Function RemoveSCGroup(ByVal index As Long) As Long
    Debug.Assert m_lngGroupCount > 0
    Dim i As Long
    ' CAUTION: This function expects that elements are removed from tail to head
    ' it then shifts all groups to the left by 1 starting at the spot where a group was deleted
    
    If index < m_lngGroupCount Then
        For i = index To m_lngGroupCount - 1
            Set m_SC_Group(i) = m_SC_Group(i + 1)
            m_SC_Group(i).GroupIndex = i
        Next
    Else
        Set m_SC_Group(m_lngGroupCount) = Nothing ' after the shift, the last item is now a double reference of 2nd to last item and must be deleted
    End If
    
    m_lngGroupCount = m_lngGroupCount - 1
    ReDim Preserve m_SC_Group(m_lngGroupCount)
    Debug.Assert m_lngGroupCount >= 0

End Function

Public Function InitSuppliers(ByRef vArr As Variant) As Long
    ' This function updates our SCGroup listings to make sure that the list of suppliers contained within each Group is valid.
    ' If there is a supplier in the vehicle but not assigned to a group, then that supplier is placed into a new group
    ' If there is a supplier NOT in the vehicle, but is in a group, that supplier is deleted AND if its the sole member of that group
    ' then the entire group is deleted
    
    ' Actually, this is somewhat insane as well (this is what happens when you code for a few hours, stop for a few days, then try to
    ' remember what it was you were thinking several days ago)  Why check that in existing profile, components in the available list match those
    ' in the SC_Groups.  Fact is, when a component is deleted or added to a vehicle, these components should be removed/added to the
    ' SC Groups in all profiles anyway.  So lets just concentrate on making this function a InitSuppliers() call which is only
    ' used on NEW profiles
    
'''''''''''''''    Insane Code commented out (leaving all this crap here just so I dont forget why this is stupid
'''''''''''''''    Dim i As Long
'''''''''''''''    Dim lngNumTotalSuppliersInVehicle As Long
'''''''''''''''    Dim lngNumGroups As Long
'''''''''''''''    Dim lngNumSuppliersInGroup As Long
'''''''''''''''
'''''''''''''''    lngNumTotalSuppliersInVehicle = UBound(vArr)
'''''''''''''''    lngNumGroups = UBound(m_SC_Group)
'''''''''''''''
'''''''''''''''    For i = 0 To lngNumGroups
'''''''''''''''        lngNumSuppliersInGroup = m_SC_Group(i).SupplierCount
'''''''''''''''        For j = 0 To lngNumSuppliersInGroup
'''''''''''''''            If m_SC_Group(i).Supplier(j) = vArr(H) Then bFound = True
'''''''''''''''        Next
'''''''''''''''    Next


    Dim i As Long
    Dim lngNumTotalSuppliersInVehicle As Long
    Dim oGroup As clsSupplyConsumeGroup
    Dim s As String
    
    lngNumTotalSuppliersInVehicle = UBound(vArr)
   
    If (lngNumTotalSuppliersInVehicle = 1) And (vArr(1) = "") Then Exit Function
    
    ReDim m_SC_Group(lngNumTotalSuppliersInVehicle)
    For i = 1 To lngNumTotalSuppliersInVehicle
        Set oGroup = New clsSupplyConsumeGroup
        Set m_SC_Group(i) = oGroup
        m_SC_Group(i).GroupIndex = i
        m_lngGroupCount = i
        Set oGroup = Nothing
        s = vArr(i)
        m_SC_Group(i).AddSupplier s
    Next
    
    InitSuppliers = True
    Exit Function
err:
InfoPrint 0, "clsProfilePower:InitSuppliers() -- Error #" & err.Number & " -- " & err.Description
InitSuppliers = False
End Function

Public Function InitConsumers(ByRef vArr As Variant)
    Dim i As Long
    Dim lngNumTotalConsumersInVehicle As Long
    ' when a profile is just created, all consumers are thrown into the "available" list and none
    ' belong to any SC_Groups at this point
    
    ' with an existing profile, when a new component is added, it also goes to "available" list and is
    ' not assigned to any SC_Group
    
    ' with an existing profile, when a component is deleted, it must be deleted from any SC_Group or Available list
    
    ' the single biggest conditions is, with an existing profile, sorting the available list from those already assigned to a group
    ' in the profile.
    ' to simplify this, we could save the local unAssignedConsumerList... then we rely on accurate node adding/removal tracking to rebuild these
    ' profiles correctly from saved .veh files.  As long as the files in the local unassignedconsumerlist as well as those assigned to Groups
    ' are always accurate, there will never be a need to sort the "Total Available Consumers" list to produce an accuracte UnAssignedConsumerList.
    
    ' in new profile, entire array of consuming components goto UnAssignedConsumerList
'    m_sUnAssignedConsumerList = vArr
'    lngNumTotalConsumersInVehicle = UBound(m_sUnAssignedConsumerList)
'
'    InitConsumers = True
    Exit Function
err:
InfoPrint 0, "clsProfilePower:InitConsumers() -- Error #" & err.Number & " -- " & err.Description
InitConsumers = False
End Function



Public Property Get Key() As String
    Key = sKey
End Property

Public Property Let Key(ByRef sData As String)
    sKey = sData
End Property

Public Property Let Description(ByVal vdata As String)
    mvarDescription = vdata
End Property

Public Property Get Description() As String
    Description = mvarDescription
End Property

Private Sub Class_Initialize()
    ReDim m_SC_Group(0)

    Set m_SC_Group(0) = New clsSupplyConsumeGroup
End Sub

Private Sub Class_Terminate()
    Dim i As Long
    For i = 0 To UBound(m_SC_Group)
        Set m_SC_Group(i) = Nothing
    Next
    
    Set m_TreeX = Nothing
End Sub





