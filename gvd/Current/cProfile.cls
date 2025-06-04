VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' weapon profiles only contain "suppliers" there are no consuming child devices to assign.
'--------------------------------------------------------------
' tracks active profile for mode switching (TODO: This crap shouldnt even be here... its a total GUI thing...)
'Private m_sActiveWeaponProfile As String
'Private m_sActiveBatteryProfile As String
'Private m_sActiveFuelProfile As String
'Private m_sActivePerformanceProfile As String
Private m_lngActiveProfileType As Long
Private m_lngActiveCheckListType As Long
Private m_sActiveCheckList As String
Private m_sActiveProfile As String
Private m_sActiveComponent As String

' link objects
' What is a Profile object?  A profile is a user defined scenario which describes
' how their vehicle manages and assigns power and fuel to devices which supply or use power and fuel.  In the
' case of weapons, it defines which weapons will work together in a group.  In the case of performance, it defines
' which thrust generating devices and which options shall be used to determine the performance of the vehicle.  A user
' can create as many profiles for all of the above as they wish.

Public Property Get ActiveProfile() As String
    ActiveProfile = m_sActiveProfile
End Property

Public Property Let ActiveProfile(ByRef s As String)
    m_sActiveProfile = s
End Property

Public Property Get ActiveProfileType() As Long
    ActiveProfileType = m_lngActiveProfileType
End Property

Public Property Let ActiveProfileType(ByVal l As Long)
    m_lngActiveProfileType = l
End Property

Public Property Get ActiveCheckList() As String
    ActiveCheckList = m_sActiveCheckList
End Property

Public Property Let ActiveCheckList(ByRef s As String)
    m_sActiveCheckList = s
End Property

Public Property Get ActiveCheckListType() As Long
    ActiveCheckListType = m_lngActiveCheckListType
End Property

Public Property Let ActiveCheckListType(ByVal l As Long)
    m_lngActiveCheckListType = l
End Property
'Public Property Get ActiveWeaponProfile() As String
'    ActiveWeaponProfile = m_sActiveWeaponProfile
'End Property
'
'Friend Property Let ActiveWeaponProfile(ByRef s As String)
'    m_sActiveWeaponProfile = s
'End Property
'
'Public Property Get ActiveBatteryProfile() As String
'    ActiveBatteryProfile = m_sActiveBatteryProfile
'End Property
'
'Friend Property Let ActiveBatteryProfile(ByRef s As String)
'    m_sActiveBatteryProfile = s
'End Property
'
'Public Property Get ActivePerformanceProfile() As String
'    ActivePerformanceProfile = m_sActivePerformanceProfile
'End Property
'
'Friend Property Let ActivePerformanceProfile(ByRef s As String)
'    m_sActivePerformanceProfile = s
'End Property
'
'Public Property Get ActiveFuelProfile() As String
'    ActiveFuelProfile = m_sActiveFuelProfile
'End Property
'
'Friend Property Let ActiveFuelProfile(ByRef s As String)
'    m_sActiveFuelProfile = s
'End Property


'todo: must implement cIPersist and cINode and cIDisplay
'-------------
Public Function AddProfile(ByRef oTreeX As TreeX, ByRef sProfileName As String, ByVal lngType As Long, ByRef sName As String) As Long
    Dim O As clsProfilePower
    Dim suppliers() As String
    Dim consumers() As String
    
    On Error GoTo err
    Set O = New clsProfilePower
    O.Key = sProfileName
    O.Description = sName
    
    Profiles.Add O, sProfileName
    Set O = Nothing
    
    If lngType = FUEL_PROFILE Then
        ' retreive list  of all suppliers in the vehicle
        suppliers = KeyManager.GetCurrentFuelStorageKeys
        ' retreive list of all consumers in the vehicle
        consumers = KeyManager.GetCurrentFuelUsingSystemKeys
    Else
        Debug.Assert lngType = POWER_PROFILE
        ' retreive list  of all suppliers in the vehicle
        suppliers = KeyManager.GetCurrentPowerSystemKeys
        ' retreive list of all consumers in the vehicle
        consumers = KeyManager.GetCurrentPowerConsumptionKeys
    End If
    
    ' pass the list of suppliers to the Profile's .InitSuppliers (suppliers)
    Call Profiles(sProfileName).InitSuppliers(suppliers)
    
    ' associate the treex to this profile
    Set Profiles(sProfileName).Tree = oTreeX
    Profiles(sProfileName).ProfileType = lngType
    Profiles(sProfileName).Show

    ' pass the list of consumers to the new PowerProfile's .InitConsumers(consumers)
     Profiles(sProfileName).UnAssignedConsumerList = consumers
    
    AddProfile = True
Exit Function

err:
    AddProfile = False
    Debug.Print "clsVehicle.AddProfile - ERROR# " & err.Number & " "; err.Description
End Function

Public Function AddConsumerToAllProfiles(ByRef sKey As String, ByVal lngProfileType As Long)
    ' When a FuelSystem component is added to the vehicle, this is called
    ' to add this components key reference to all fuel profiles that exist
    Dim O As clsProfilePower
    
    For Each O In Profiles
        If O.ProfileType = lngProfileType Then O.AddConsumer sKey
    Next
End Function

Public Function RemoveConsumerFromAllProfiles(ByRef sKey As String, ByVal lngProfileType As Long)
    ' When a component is deleted from the vehicle, this is called to
    ' remove this components key reference from all fuel profiles that exist
    Dim O As clsProfilePower
    
    For Each O In Profiles
        If O.ProfileType = lngProfileType Then O.RemoveConsumer sKey
    Next
End Function

Public Function AddSupplierToAllProfiles(ByRef sKey As String, ByVal lngProfileType As Long)
    ' When a Fuel Supplier (tank) component is added to the vehicle, this is called
    ' to add this components key reference to all power profiles that exist
    Dim O As clsProfilePower
    
    For Each O In Profiles
        If O.ProfileType = lngProfileType Then O.AddPowerSystem sKey
    Next
End Function

Public Function RemoveSupplierFromAllProfiles(ByRef sKey As String, ByVal lngProfileType As Long)
    ' When a tank is deleted from the vehicle, this is called to
    ' remove this components key reference from all power profiles that exist
    Dim O As clsProfilePower
    
    'NOTE: Checking for orphaned children is done in the RemoveEmptySCGroup function inside of clsPowerProfiles
    For Each O In Profiles
        If O.ProfileType = lngProfileType Then O.RemovePowerSystem sKey
    Next
End Function



Private Sub Class_Initialize()
    ' todo: init any arrays
End Sub

Private Sub Class_Terminate()
' todo: terminate all profiles in here
End Sub
