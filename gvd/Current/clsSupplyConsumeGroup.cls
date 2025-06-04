VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSupplyConsumeGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private m_sSupplier() As String 'holds keys of suppliers in this particular group
Private m_sConsumer() As String ' holds keys of consumers in this particular group
Private m_lngGroupIndex As Long

Private m_lngSupplierCount As Long
Private m_lngConsumerCount As Long

Private Sub Class_Initialize()
    m_lngConsumerCount = 0
    m_lngSupplierCount = 0
End Sub

Private Sub Class_Terminate()

    Erase m_sSupplier
    Erase m_sConsumer
End Sub

Public Property Let GroupIndex(ByVal L As Long)
    m_lngGroupIndex = L
End Property

Public Sub Show(ByRef O As Object)
    ' check that there is at least one Supplier so we know if a Group node is needed
    Dim hGroup As Long
    Dim hChildGroup As Long
    Dim hSupplier As Long
    Dim hConsumer As Long
    Dim oTree As TreeX
    Dim i As Long
    Set oTree = O
    
    
        
        hGroup = oTree.AddItem(GROUP_NAME)
        oTree.ItemData(hGroup) = m_lngGroupIndex
        hChildGroup = oTree.AddItem(CHILD_GROUP_NAME, hGroup, 0)
        oTree.ItemData(hChildGroup) = m_lngGroupIndex
    
    If m_lngSupplierCount >= 1 Then
        For i = 1 To UBound(m_sSupplier)
            hSupplier = oTree.AddItem(Veh.Components(m_sSupplier(i)).CustomDescription, hGroup)
            oTree.ItemData(hSupplier) = i
        Next
        oTree.ExpandItem(hGroup) = True
    End If
    
    On Error GoTo err
    
    If m_lngConsumerCount >= 1 Then
        For i = 1 To UBound(m_sConsumer)
            hConsumer = oTree.AddItem(Veh.Components(m_sConsumer(i)).CustomDescription, hChildGroup)
            oTree.ItemData(hConsumer) = i
        Next
        oTree.ExpandItem(hChildGroup) = True
    End If
err:
    Set oTree = Nothing
End Sub

Public Function Consumer(ByVal index As Long) As String
    Consumer = m_sConsumer(index)
End Function
Public Function Supplier(ByVal index As Long) As String
    Supplier = m_sSupplier(index)
End Function

Public Function ConsumerCount() As Long
    On Error GoTo err
    ConsumerCount = UBound(m_sConsumer)
    Exit Function
err:
    ConsumerCount = 0
End Function
Public Function SupplierCount() As Long
    SupplierCount = UBound(m_sSupplier)
End Function

Public Function AddConsumer(ByVal s As String) As Long
' function accepts the Consumer's key value
    m_lngConsumerCount = m_lngConsumerCount + 1
    ReDim Preserve m_sConsumer(m_lngConsumerCount)
    
    m_sConsumer(m_lngConsumerCount) = s
    AddConsumer = True

End Function

Public Function RemoveConsumer(ByRef s As String) As Long
    Dim index As Long
    Dim bFound As Boolean
    
    For index = 1 To m_lngConsumerCount
        If m_sConsumer(index) = s Then
            bFound = True
            Exit For
        End If
    Next

    If Not bFound Then Exit Function
    Call RemoveConsumerByIndex(index)
End Function
Public Function RemoveConsumerByIndex(ByVal index As Long) As Long
    Dim i As Long
    
    If m_lngConsumerCount <> index Then
        For i = index To m_lngConsumerCount - 1
            m_sConsumer(i) = m_sConsumer(i + 1)
        Next
    End If
    
    m_lngConsumerCount = m_lngConsumerCount - 1
    ReDim Preserve m_sConsumer(m_lngConsumerCount)
    Debug.Assert m_lngConsumerCount >= 0

End Function

Public Function AddSupplier(ByRef s As String) As Long
    m_lngSupplierCount = m_lngSupplierCount + 1
    ReDim Preserve m_sSupplier(m_lngSupplierCount)
    
    m_sSupplier(m_lngSupplierCount) = s
    AddSupplier = True
End Function

Public Function RemoveSupplier(ByRef s As String) As Long
    Dim index As Long
    Dim bFound As Boolean
    
    For index = 1 To m_lngSupplierCount
        If m_sSupplier(index) = s Then
            bFound = True
            Exit For
        End If
    Next
    
    If Not bFound Then Exit Function
    Call RemoveSupplierByIndex(index)

End Function

Public Function RemoveSupplierByIndex(ByVal lngIndex As Long)
    Dim i As Long
    
    If m_lngSupplierCount <> lngIndex Then
        For i = lngIndex To m_lngSupplierCount - 1
            m_sSupplier(i) = m_sSupplier(i + 1)
        Next
        ' copy memory cant work here as is unless I loop... im working with variable length string arrays and not a
        ' continuous block of memory
        'CopyMemory m_sSupplier(lngIndex), m_sSupplier(lngIndex + 1), m_lngSupplierCount - lngIndex
    End If
    m_lngSupplierCount = m_lngSupplierCount - 1
    ReDim Preserve m_sSupplier(m_lngSupplierCount)
    Debug.Assert m_lngSupplierCount >= 0
End Function
