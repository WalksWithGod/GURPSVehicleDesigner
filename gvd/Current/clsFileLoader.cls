VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Implements tli.CustomFilter
Private m_IFace As tli.InterfaceInfo


'//////////////////////////////
Public Function GetProperties(objVC As Object) As Variant

Dim SR As SearchResults
Dim SI As SearchItem
Dim vc() As Variant

Dim i As Long
    
    i = 1
    
    Set m_IFace = tli.InterfaceInfoFromObject(objVC)
    
    With m_IFace.Members
        .FuncFilter = .FuncFilter Or FUNCFLAG_FNONBROWSABLE
        .VarFilter = .VarFilter Or VARFLAG_FNONBROWSABLE
        Set SR = .GetFilteredMembers
        SR.Filter Me 'Optional, but good for sample
        ReDim Preserve vc(SR.Count, SR.Count)
        On Error Resume Next
        For Each SI In SR
            
            With SI
                vc(i, 0) = .name
                vc(0, i) = tli.InvokeHook(objVC, .memberID, INVOKE_PROPERTYGET)
                If .name = "MotivePower" Then
                    Debug.Print vc(0, i)
                End If
                i = i + 1
            End With
        Next
        Err.Clear
    End With

    
GetProperties = vc
End Function

Public Function LetProperties(veh As Object, memberID As String, value As Variant) As Boolean

    On Error GoTo errorhandler
    
    tli.InvokeHook veh, memberID, INVOKE_PROPERTYPUT, value
    
    Exit Function
    
errorhandler:
    If Err.Number = 13 Then   'type mismatch
        On Error Resume Next
        tli.InvokeHook veh, memberID, INVOKE_PROPERTYPUT, Val(value)
    End If
    Err.Clear
   
    
End Function

'All of this filtering is actually optional since calling InvokeHook with
'PROPERTYGET will return an error.  However, this does demonstrate how to
'apply a custom filter.
Private Sub CustomFilter_Visit(ByVal item As tli.SearchItem, Action As tli.TliCustomFilterAction)

    If item.InvokeKinds And INVOKE_PROPERTYGET Then
        With m_IFace.GetMember(item.memberID)
            If .Parameters.Count Then
                With .Parameters
                    If .DefaultCount = 0 Then
                        If .OptionalCount = 0 Then
                            If .Count Then Action = tliCfaDelete
                        End If
                    End If
                End With
            Else
                With ResolveVarTypeInfo(.ReturnType)
                    Select Case .VarType
                        Case VT_EMPTY
                            If .TypeInfo.TypeKind <> TKIND_ENUM Then Action = tliCfaDelete
                        Case VT_DISPATCH, VT_UNKNOWN
                            Action = tliCfaDelete
                    End Select
                End With
            End If
        End With
    Else
        Action = tliCfaDelete
    End If
End Sub

Private Function ResolveVarTypeInfo(ByVal VTI As VarTypeInfo) As VarTypeInfo
    If VTI.VarType = VT_EMPTY Then 'Indicates a TypeLib defined type.
        With VTI.TypeInfo
            If .TypeKind = TKIND_ALIAS Then
                Set ResolveVarTypeInfo = ResolveVarTypeInfo(.ResolvedType)
                Exit Function
            End If
        End With
    End If
    Set ResolveVarTypeInfo = VTI
End Function

