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
vbwProfiler.vbwProcIn 29

Dim SR As SearchResults
Dim SI As SearchItem
Dim vc() As Variant

Dim i As Long

vbwProfiler.vbwExecuteLine 944
    i = 1

vbwProfiler.vbwExecuteLine 945
    Set m_IFace = tli.InterfaceInfoFromObject(objVC)

vbwProfiler.vbwExecuteLine 946
    With m_IFace.Members
vbwProfiler.vbwExecuteLine 947
        .FuncFilter = .FuncFilter Or FUNCFLAG_FNONBROWSABLE
vbwProfiler.vbwExecuteLine 948
        .VarFilter = .VarFilter Or VARFLAG_FNONBROWSABLE
vbwProfiler.vbwExecuteLine 949
        Set SR = .GetFilteredMembers
vbwProfiler.vbwExecuteLine 950
        SR.Filter Me 'Optional, but good for sample
vbwProfiler.vbwExecuteLine 951
        ReDim Preserve vc(SR.Count, SR.Count)
vbwProfiler.vbwExecuteLine 952
        On Error Resume Next
vbwProfiler.vbwExecuteLine 953
        For Each SI In SR

vbwProfiler.vbwExecuteLine 954
            With SI
vbwProfiler.vbwExecuteLine 955
                vc(i, 0) = .name
vbwProfiler.vbwExecuteLine 956
                vc(0, i) = tli.InvokeHook(objVC, .memberID, INVOKE_PROPERTYGET)
vbwProfiler.vbwExecuteLine 957
                If .name = "MotivePower" Then
vbwProfiler.vbwExecuteLine 958
                    Debug.Print vc(0, i)
                End If
vbwProfiler.vbwExecuteLine 959 'B
vbwProfiler.vbwExecuteLine 960
                i = i + 1
vbwProfiler.vbwExecuteLine 961
            End With
vbwProfiler.vbwExecuteLine 962
        Next
vbwProfiler.vbwExecuteLine 963
        Err.Clear
vbwProfiler.vbwExecuteLine 964
    End With


vbwProfiler.vbwExecuteLine 965
GetProperties = vc
vbwProfiler.vbwProcOut 29
vbwProfiler.vbwExecuteLine 966
End Function

Public Function LetProperties(veh As Object, memberID As String, value As Variant) As Boolean
vbwProfiler.vbwProcIn 30

vbwProfiler.vbwExecuteLine 967
    On Error GoTo errorhandler

vbwProfiler.vbwExecuteLine 968
    tli.InvokeHook veh, memberID, INVOKE_PROPERTYPUT, value

vbwProfiler.vbwProcOut 30
vbwProfiler.vbwExecuteLine 969
    Exit Function

errorhandler:
vbwProfiler.vbwExecuteLine 970
    If Err.Number = 13 Then   'type mismatch
vbwProfiler.vbwExecuteLine 971
        On Error Resume Next
vbwProfiler.vbwExecuteLine 972
        tli.InvokeHook veh, memberID, INVOKE_PROPERTYPUT, Val(value)
    End If
vbwProfiler.vbwExecuteLine 973 'B
vbwProfiler.vbwExecuteLine 974
    Err.Clear


vbwProfiler.vbwProcOut 30
vbwProfiler.vbwExecuteLine 975
End Function

'All of this filtering is actually optional since calling InvokeHook with
'PROPERTYGET will return an error.  However, this does demonstrate how to
'apply a custom filter.
Private Sub CustomFilter_Visit(ByVal item As tli.SearchItem, Action As tli.TliCustomFilterAction)
vbwProfiler.vbwProcIn 31

vbwProfiler.vbwExecuteLine 976
    If item.InvokeKinds And INVOKE_PROPERTYGET Then
vbwProfiler.vbwExecuteLine 977
        With m_IFace.GetMember(item.memberID)
vbwProfiler.vbwExecuteLine 978
            If .Parameters.Count Then
vbwProfiler.vbwExecuteLine 979
                With .Parameters
vbwProfiler.vbwExecuteLine 980
                    If .DefaultCount = 0 Then
vbwProfiler.vbwExecuteLine 981
                        If .OptionalCount = 0 Then
vbwProfiler.vbwExecuteLine 982
                            If .Count Then
vbwProfiler.vbwExecuteLine 983
                                 Action = tliCfaDelete
                            End If
vbwProfiler.vbwExecuteLine 984 'B
                        End If
vbwProfiler.vbwExecuteLine 985 'B
                    End If
vbwProfiler.vbwExecuteLine 986 'B
vbwProfiler.vbwExecuteLine 987
                End With
            Else
vbwProfiler.vbwExecuteLine 988 'B
vbwProfiler.vbwExecuteLine 989
                With ResolveVarTypeInfo(.ReturnType)
vbwProfiler.vbwExecuteLine 990
                    Select Case .VarType
'vbwLine 991:                        Case VT_EMPTY
                        Case IIf(vbwProfiler.vbwExecuteLine(991), VBWPROFILER_EMPTY, _
        VT_EMPTY)
vbwProfiler.vbwExecuteLine 992
                            If .TypeInfo.TypeKind <> TKIND_ENUM Then
vbwProfiler.vbwExecuteLine 993
                                 Action = tliCfaDelete
                            End If
vbwProfiler.vbwExecuteLine 994 'B
'vbwLine 995:                        Case VT_DISPATCH, VT_UNKNOWN
                        Case IIf(vbwProfiler.vbwExecuteLine(995), VBWPROFILER_EMPTY, _
        VT_DISPATCH), VT_UNKNOWN
vbwProfiler.vbwExecuteLine 996
                            Action = tliCfaDelete
                    End Select
vbwProfiler.vbwExecuteLine 997 'B
vbwProfiler.vbwExecuteLine 998
                End With
            End If
vbwProfiler.vbwExecuteLine 999 'B
vbwProfiler.vbwExecuteLine 1000
        End With
    Else
vbwProfiler.vbwExecuteLine 1001 'B
vbwProfiler.vbwExecuteLine 1002
        Action = tliCfaDelete
    End If
vbwProfiler.vbwExecuteLine 1003 'B
vbwProfiler.vbwProcOut 31
vbwProfiler.vbwExecuteLine 1004
End Sub

Private Function ResolveVarTypeInfo(ByVal VTI As VarTypeInfo) As VarTypeInfo
vbwProfiler.vbwProcIn 32
vbwProfiler.vbwExecuteLine 1005
    If VTI.VarType = VT_EMPTY Then 'Indicates a TypeLib defined type.
vbwProfiler.vbwExecuteLine 1006
        With VTI.TypeInfo
vbwProfiler.vbwExecuteLine 1007
            If .TypeKind = TKIND_ALIAS Then
vbwProfiler.vbwExecuteLine 1008
                Set ResolveVarTypeInfo = ResolveVarTypeInfo(.ResolvedType)
vbwProfiler.vbwProcOut 32
vbwProfiler.vbwExecuteLine 1009
                Exit Function
            End If
vbwProfiler.vbwExecuteLine 1010 'B
vbwProfiler.vbwExecuteLine 1011
        End With
    End If
vbwProfiler.vbwExecuteLine 1012 'B
vbwProfiler.vbwExecuteLine 1013
    Set ResolveVarTypeInfo = VTI
vbwProfiler.vbwProcOut 32
vbwProfiler.vbwExecuteLine 1014
End Function


