Option Strict Off
Option Explicit On
Friend Class clsFileLoader
	Implements tli._CustomFilter
	
	Private m_IFace As tli.InterfaceInfo
	
	
	'//////////////////////////////
	Public Function GetProperties(ByRef objVC As Object) As Object
		Dim INVOKE_PROPERTYGET As Object
		Dim VARFLAG_FNONBROWSABLE As Object
		Dim FUNCFLAG_FNONBROWSABLE As Object
		Dim tli As Object
		
		'UPGRADE_ISSUE: SearchResults object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim SR As SearchResults
		'UPGRADE_ISSUE: SearchItem object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim SI As SearchItem
		Dim vc() As Object
		
		Dim i As Integer
		
		i = 1
		
		'UPGRADE_WARNING: Couldn't resolve default property of object tli.InterfaceInfoFromObject. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		m_IFace = tli.InterfaceInfoFromObject(objVC)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_IFace.Members. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With m_IFace.Members
			'UPGRADE_WARNING: Couldn't resolve default property of object m_IFace.Members. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.FuncFilter = .FuncFilter Or FUNCFLAG_FNONBROWSABLE
			'UPGRADE_WARNING: Couldn't resolve default property of object m_IFace.Members. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.VarFilter = .VarFilter Or VARFLAG_FNONBROWSABLE
			'UPGRADE_WARNING: Couldn't resolve default property of object m_IFace.Members. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SR = .GetFilteredMembers
			'UPGRADE_WARNING: Couldn't resolve default property of object SR.Filter. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SR.Filter(Me) 'Optional, but good for sample
			'UPGRADE_WARNING: Couldn't resolve default property of object SR.Count. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			ReDim Preserve vc(SR.Count, SR.Count)
			On Error Resume Next
			For	Each SI In SR
				
				With SI
					'UPGRADE_WARNING: Couldn't resolve default property of object SI.name. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object vc(i, 0). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					vc(i, 0) = .name
					'UPGRADE_WARNING: Couldn't resolve default property of object SI.memberID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object tli.InvokeHook. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object vc(0, i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					vc(0, i) = tli.InvokeHook(objVC, .memberID, INVOKE_PROPERTYGET)
					'UPGRADE_WARNING: Couldn't resolve default property of object SI.name. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .name = "MotivePower" Then
						Debug.Print(vc(0, i))
					End If
					i = i + 1
				End With
			Next SI
			Err.Clear()
		End With
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object GetProperties. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GetProperties = VB6.CopyArray(vc)
	End Function
	
	Public Function LetProperties(ByRef veh As Object, ByRef memberID As String, ByRef value As Object) As Boolean
		Dim INVOKE_PROPERTYPUT As Object
		Dim tli As Object
		
		On Error GoTo errorhandler
		
		'UPGRADE_WARNING: Couldn't resolve default property of object tli.InvokeHook. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		tli.InvokeHook(veh, memberID, INVOKE_PROPERTYPUT, value)
		
		Exit Function
		
errorhandler: 
		If Err.Number = 13 Then 'type mismatch
			On Error Resume Next
			'UPGRADE_WARNING: Couldn't resolve default property of object value. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object tli.InvokeHook. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			tli.InvokeHook(veh, memberID, INVOKE_PROPERTYPUT, Val(value))
		End If
		Err.Clear()
		
		
	End Function
	
	'All of this filtering is actually optional since calling InvokeHook with
	'PROPERTYGET will return an error.  However, this does demonstrate how to
	'apply a custom filter.
	Private Sub CustomFilter_Visit(ByVal item As tli.SearchItem, ByRef Action As tli.TliCustomFilterAction)
		Dim INVOKE_PROPERTYGET As Object
		Dim VT_UNKNOWN As Object
		Dim VT_DISPATCH As Object
		Dim TKIND_ENUM As Object
		Dim VT_EMPTY As Object
		Dim tliCfaDelete As Object
		Dim tli As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object item.InvokeKinds. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If item.InvokeKinds And INVOKE_PROPERTYGET Then
			'UPGRADE_WARNING: Couldn't resolve default property of object item.memberID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object m_IFace.GetMember. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			With m_IFace.GetMember(item.memberID)
				'UPGRADE_WARNING: Couldn't resolve default property of object item.memberID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object m_IFace.GetMember. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If .Parameters.Count Then
					'UPGRADE_WARNING: Couldn't resolve default property of object item.memberID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object m_IFace.GetMember. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					With .Parameters
						'UPGRADE_WARNING: Couldn't resolve default property of object item.memberID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object m_IFace.GetMember. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If .DefaultCount = 0 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object item.memberID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object m_IFace.GetMember. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If .OptionalCount = 0 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object item.memberID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object m_IFace.GetMember. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object Action. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object tliCfaDelete. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If .Count Then Action = tliCfaDelete
							End If
						End If
					End With
				Else
					'UPGRADE_WARNING: Couldn't resolve default property of object item.memberID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object m_IFace.GetMember. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					With ResolveVarTypeInfo(.ReturnType)
						'UPGRADE_WARNING: Couldn't resolve default property of object VT_UNKNOWN. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object VT_DISPATCH. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object VT_EMPTY. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object item.memberID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object m_IFace.GetMember. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object ResolveVarTypeInfo(m_IFace.GetMember(item.memberID).ReturnType).VarType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						Select Case .VarType
							Case VT_EMPTY
								'UPGRADE_WARNING: Couldn't resolve default property of object item.memberID. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object m_IFace.GetMember. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object ResolveVarTypeInfo(m_IFace.GetMember(item.memberID).ReturnType).TypeInfo. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object Action. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object tliCfaDelete. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If .TypeInfo.TypeKind <> TKIND_ENUM Then Action = tliCfaDelete
							Case VT_DISPATCH, VT_UNKNOWN
								'UPGRADE_WARNING: Couldn't resolve default property of object Action. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object tliCfaDelete. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								Action = tliCfaDelete
						End Select
					End With
				End If
			End With
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object Action. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object tliCfaDelete. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Action = tliCfaDelete
		End If
	End Sub
	
	Private Function ResolveVarTypeInfo(ByVal VTI As VarTypeInfo) As VarTypeInfo
		Dim VT_EMPTY As Object
		Dim TKIND_ALIAS As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object VTI.VarType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If VTI.VarType = VT_EMPTY Then 'Indicates a TypeLib defined type.
			'UPGRADE_WARNING: Couldn't resolve default property of object VTI.TypeInfo. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			With VTI.TypeInfo
				'UPGRADE_WARNING: Couldn't resolve default property of object VTI.TypeInfo. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If .TypeKind = TKIND_ALIAS Then
					'UPGRADE_WARNING: Couldn't resolve default property of object VTI.TypeInfo. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ResolveVarTypeInfo = ResolveVarTypeInfo(.ResolvedType)
					Exit Function
				End If
			End With
		End If
		ResolveVarTypeInfo = VTI
	End Function
End Class