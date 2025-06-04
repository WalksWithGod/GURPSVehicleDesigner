VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements cIPersist
Implements cIDisplay

Private m_sName As String
Private m_sClassname As String

Private m_lngModifierCount As Long
Private m_oProperties() As vehicles.cpropertyitem
Private m_lngCurrentPropItem As Long
Private m_lngPropCount As Long

Private m_oAuthor As vehicles.cAuthor
Private m_oVersion As vehicles.cVersion

Private m_sIcon As String 'path
Private m_lngMatrix() As Single


Public Property Get propertyCount() As Long
    propertyCount = m_lngPropertyCount
End Property

Public Property Get className() As String
    className = m_sClassname
End Property

'//cIDisplay Implemented Properties and Functions
Private Function cIDisplay_getFirstPropertyItem() As cpropertyitem
    If Not m_oProperties(0) Is Nothing Then
        Set cIDisplay_getFirstPropertyItem = m_oProperties(0)
        m_lngCurrentPropItem = 0
    End If
End Function

Private Function cIDisplay_getNextPropertyItem() As cpropertyitem
    m_lngCurrentPropItem = m_lngCurrentPropItem + 1
    If m_lngCurrentPropItem <= m_lngPropCount - 1 Then
        If Not m_oProperties(m_lngCurrentPropItem) Is Nothing Then
            Set cIDisplay_getNextPropertyItem = m_oProperties(m_lngCurrentPropItem)
        End If
    Else
        m_lngCurrentPropItem = m_lngCurrentPropItem - 1
        Debug.Print TypeName(Me) & ":cIDisplay:getNextPropertyItem -- nextItem exceeds Property Count."
    End If
End Function
Private Function cIDisplay_getPropertyItemByIndex(ByVal iIndex As Long) As cpropertyitem
    On Error Resume Next
    Set cIDisplay_getPropertyItemByIndex = m_oProperties(iIndex)
End Function

'-----------------

Private Property Get cIPersist_Classname() As String
End Property

Private Property Get cIPersist_GUID() As String
End Property

Private Sub cIPersist_LoadProperties(ByVal op As PersistenceManager.clsObjProperties, ByVal iMode As PersistenceManager.GVD_XML_TYPE)
    Dim i As Long
    Dim o As Object
    
    Set m_oAuthor = op.Load("author")
    Set m_oVersion = op.Load("version")
    m_sClassname = op.Load("classname")
    
    m_lngPropCount = op.Load("propertycount")
    
    If m_lngPropCount > 0 Then
        ReDim m_oProperties(m_lngPropCount - 1)
        For i = 0 To m_lngPropCount - 1
            Set m_oProperties(i) = op.Load("property" & i)
        Next
    End If
        

End Sub

Private Sub cIPersist_StoreProperties(ByVal op As PersistenceManager.clsObjProperties)
End Sub

Private Sub Class_Terminate()
    Dim i As Long
    For i = 0 To m_lngPropCount - 1
        Set m_oProperties(i) = Nothing
    Next
End Sub
