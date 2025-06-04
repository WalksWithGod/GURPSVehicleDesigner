Option Strict Off
Option Explicit On
Module modProperties
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Sub CopyMemory Lib "kernel32"  Alias "RtlMoveMemory"(ByRef hpvDest As Any, ByRef hpvSource As Any, ByVal cbCopy As Integer)
	Private Const PROPERTY_HEADER As String = "HEADER"
	'constants for the PropertyList datatypes
	'todo: the CmpEdit will need to use these too... is this bas shared or should we rip out the constants and put them in a shareable module?
	Private Const wdBool As Short = 1
	Private Const wdColor As Short = 4
	Private Const wdCurrency As Short = 9
	Private Const wdDate As Short = 6
	Private Const wdDefault As Short = -1
	Private Const wdDouble As Short = 10
	Private Const wdFile As Short = 5
	Private Const wdFont As Short = 2
	Private Const wdList As Short = 3
	Private Const wdNumber As Short = 8
	Private Const wdObject As Short = 11
	Private Const wdPicture As Short = 7
	Private Const wdText As Short = 0
	Private Const wdHeader As Short = -99
	
	Function formatCaption(ByVal lngDatatType As Integer, ByVal lngUnitType As Integer, ByRef s As String) As String
		s = " " & s
		formatCaption = s
	End Function
	
	Function formatValue(ByVal lngDatatype As Integer, ByVal lngUnitType As Integer, ByRef v As Object) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object v. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object formatValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		formatValue = v
	End Function
	
	Sub AddPCLproperty(ByRef vValue As Object, ByRef oPropItem As cPropertyItem)
		Dim frmDesigner As Object
		Dim lngNewIndex As Integer
		Dim lngDatatype As Integer
		Dim lngUnitType As Integer
		Dim sCaption As String
		Dim dblValue As Double
		Dim vItem As Object
		Dim i As Integer
		'Const SEPERATOR = "___________________"
		Const SEPERATOR As String = "===================="
		
		'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lngDatatype = oPropItem.Datatype ' proplist data type (e.g. list, float, number, text, etc)
		'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.UnitType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lngUnitType = oPropItem.UnitType ' gvd unit type used by the unit converter
		'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.Caption. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCaption = oPropItem.Caption
		
		sCaption = formatCaption(lngDatatype, lngUnitType, sCaption)
		'UPGRADE_WARNING: Couldn't resolve default property of object formatValue(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		vValue = formatValue(lngDatatype, lngUnitType, vValue)
		'todo: when finally done, check out for places where i handle doubles and singles together to make sure its all correct.
		' see notes below on some things to look out for
		
		' NOTE: I use double's for most final stats, however I dont for table data since it would require double the amount of space
		' and there are ALOT of tables.  Also used in our m_UserInput() variable which gets displayed directly here in the proplist.
		' However, since the PLC1 doesnt support a "single" datatype, I must use the PLC1's double datatype
		' which unfortunately will show approximation errors.  To fix this, we must convert the single to a string first and then assign that to the double.
		' NOTE: Double to single doesnt cause these errors (as long as it doesnt overflow obviously) but single to double does.  Keep this in mind during
		' stats calculations where we might be mixing the SINGLE's in our tables with the DOUBLE's required for the final answer.
		If lngDatatype = wdDouble Then
			'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			dblValue = Val(CStr(vValue))
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			frmDesigner.PLC1.AddItem(dblValue, lngDatatype)
		ElseIf lngDatatype = wdHeader Then 
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			With frmDesigner.PLC1
				'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.AddItem(SEPERATOR, 0) ' change it back from -99 to 0 so that we can display the seperator line
				'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.ItemDisabledTextBold(.NewIndex) = True
			End With
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			frmDesigner.PLC1.AddItem(vValue, lngDatatype)
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With frmDesigner.PLC1
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lngNewIndex = .NewIndex
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.ItemData(lngNewIndex) = lngDatatype
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.CaptionString(lngNewIndex) = sCaption
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.Notes. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.DescriptionString(lngNewIndex) = oPropItem.Notes
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.ReadOnly. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.ItemDisabled(lngNewIndex) = oPropItem.ReadOnly
		End With
		
		'todo: how do we handle ammolists, guidancelists, or material/quality lists for armor?
		'      the armor layer would have to update its property item's list every time techlevel was changed.
		'      and if tech level was changed to somethign not supported by the current material, the material
		'     would default to first available at that techlevel. (e.g. cheap wood).  The only way to do this
		'     properly is use multiple matrices so that the values can be looked up based on the tech level and material.
		If lngDatatype = wdList Then
			'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.List. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object vItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			vItem = oPropItem.List
			If IsArray(vItem) Then
				For i = LBound(vItem) To UBound(vItem)
					'UPGRADE_WARNING: Couldn't resolve default property of object vItem(i). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If vItem(i) <> "" Then
						'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						frmDesigner.PLC1.AddListItem(lngNewIndex, vItem(i))
					End If
				Next 
			End If
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		vValue = Nothing
	End Sub
	
	Public Sub PropertyChanged(ByVal index As Integer)
		Dim frmDesigner As Object
		Const LNG_LENGTH As Short = 4
		'UPGRADE_ISSUE: cPropertyItem object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim oProp As cPropertyItem
		Dim oNode As _cINode
		Dim oDisplay As _cIDisplay
		Dim lptr As Integer
		Dim sClassname As String
		Dim lngInterfaceID As Integer
		Dim lngRet As Integer
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Debug.Print("Object Handle = " & frmDesigner.PLC1.Tag & " PLC1 INDEX = " & index)
		
		On Error GoTo err_Renamed
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		lptr = Val(frmDesigner.PLC1.Tag)
		'UPGRADE_WARNING: Couldn't resolve default property of object oNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(oNode, lptr, LNG_LENGTH)
		
		Dim oComponent As _cIComponent
		Dim oContainer As _cIContainer
		Dim oVehicle As cVehicle
		Dim oBuild As _cIBuild
		Dim oSurface As cSurface
		If Not oNode Is Nothing Then
			sClassname = oNode.Classname
			oDisplay = oNode
			
			' NOTE: its imperative that the order of properties in the proplist match the order of properties in the object
			oProp = oDisplay.getPropertyItemByIndex(index)
			'UPGRADE_WARNING: Couldn't resolve default property of object oProp.interfaceid. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lngInterfaceID = oProp.interfaceid
			
			'UPGRADE_WARNING: Couldn't resolve default property of object oProp.ReadOnly. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If Not oProp.ReadOnly Then
				' todo: i think this select case is "ok"... anyway to modify it or perhaps consolidate the code
				'       with modProperties:PropertiesShow() ??  Actually probably not, it uses "vbGet" and this uses
				'       vbLet.  Actually what we should do then is move this out into a seperate function
				'UPGRADE_WARNING: Couldn't resolve default property of object oProp.CallByName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Debug.Print("modProperties:PropertyChanged() -- InterfaceID = " & lngInterfaceID & " PropertyName = " & oProp.CallByName)
				Select Case lngInterfaceID
					Case INTERFACE_COMPONENT
						oComponent = oNode
						'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object oProp.CallByName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						CallByName(oComponent, oProp.CallByName, CallType.Set, frmDesigner.PLC1.value(index))
						
					Case INTERFACE_CONTAINER
						oContainer = oNode
						'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object oProp.CallByName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						CallByName(oContainer, oProp.CallByName, CallType.Set, frmDesigner.PLC1.value(index))
					Case INTERFACE_VEHICLE_DESCRIPTION
						oVehicle = oNode
						'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object oProp.CallByName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						CallByName(oVehicle.Description, oProp.CallByName, CallType.Set, frmDesigner.PLC1.value(index))
					Case INTERFACE_VEHICLE_VERSION
						'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object oProp.CallByName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						CallByName(oVehicle.version, oProp.CallByName, CallType.Set, frmDesigner.PLC1.value(index))
						
					Case INTERFACE_VEHICLE_AUTHOR
						'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object oProp.CallByName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						CallByName(oVehicle.author, oProp.CallByName, CallType.Set, frmDesigner.PLC1.value(index))
						
					Case INTERFACE_NODE
						'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object oProp.CallByName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						CallByName(oNode, oProp.CallByName, CallType.Set, frmDesigner.PLC1.value(index))
					Case INTERFACE_DISPLAY
						'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object oProp.CallByName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						CallByName(oDisplay, oProp.CallByName, CallType.Set, frmDesigner.PLC1.value(index))
					Case INTERFACE_BUILD
						oBuild = oNode
						'UPGRADE_WARNING: Couldn't resolve default property of object oProp.CallBytype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If oProp.CallBytype = 1 Then
							' since this is a function call, we have to modify the name to append "get"
							' also, if its a wdList, we need to pass the subscript of the selected list item
							'UPGRADE_WARNING: Couldn't resolve default property of object oProp.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If oProp.Datatype = wdList Then
								'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object oProp.getSelectionIndexFromValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								lngRet = oProp.getSelectionIndexFromValue(frmDesigner.PLC1.value(index))
								If lngRet <> -1 Then
									'UPGRADE_WARNING: Couldn't resolve default property of object oProp.Subscript. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object oProp.CallByName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									CallByName(oBuild, "set" & oProp.CallByName, CallType.Method, oProp.Subscript, lngRet)
								Else
									' an error and there shouldnt be any
									MsgBox("modProperties:PropertyChanged() -- Invalid wdList subscript")
								End If
								' else its userinput and we pass the actual value
								'UPGRADE_WARNING: Couldn't resolve default property of object oProp.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							ElseIf oProp.Datatype = wdDouble Then 
								'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object oProp.Subscript. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object oProp.CallByName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								CallByName(oBuild, "set" & oProp.CallByName, CallType.Method, oProp.Subscript, frmDesigner.PLC1.value(index))
							Else ' we should never reach here
								
							End If
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object oProp.CallByName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							CallByName(oBuild, oProp.CallByName, CallType.Set, frmDesigner.PLC1.value(index))
						End If
					Case INTERFACE_SURFACE
						oSurface = oNode
						'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object oProp.getSelectionIndexFromValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						lngRet = oProp.getSelectionIndexFromValue(frmDesigner.PLC1.value(index))
						If lngRet <> -1 Then
							'UPGRADE_WARNING: Couldn't resolve default property of object oProp.CallByName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							CallByName(oSurface, oProp.CallByName, CallType.Set, lngRet)
						Else
							' an error and there shouldnt be any
							MsgBox("modProperties:PropertyChanged() -- Invalid wdList subscript")
						End If
						
					Case Else
						Debug.Print("modProperties:PropertyChanged() -- Class Interface Not Supported.")
				End Select
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object oProp.Caption. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				modHelper.InfoPrint(1, "The property '" & oProp.Caption & "' is Read Only.")
			End If
			
			'UPGRADE_NOTE: Object oProp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oProp = Nothing
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object oNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(oNode, 0, LNG_LENGTH)
		'UPGRADE_NOTE: Object oDisplay may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oDisplay = Nothing
		System.Windows.Forms.Application.DoEvents()
		'    UpdateVehicle todo: all of these may be obosolete under new code base EXCEPT when user hits F5 or when forcing recalc after laoding saved vehicle
		
		'//place the cell back into its original spot. Todo: why is this needed?
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.PLC1.ListIndex = index
		p_bChangedFlag = True ' JAW 2000.05.07 change has been made in vehicle/component
		Exit Sub
err_Renamed: 
		Debug.Print("modProperties:PropertyChanged() -- Error #" & Err.Number & " " & Err.Description)
		If Not oNode Is Nothing Then
			'UPGRADE_WARNING: Couldn't resolve default property of object oNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(oNode, 0, LNG_LENGTH)
			'UPGRADE_NOTE: Object oDisplay may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oDisplay = Nothing
			'UPGRADE_NOTE: Object oProp may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oProp = Nothing
		End If
		
		'    Todo: still need to properly handle these types.... perhaps new datatype like wdNote (signifies text greater than X chars?)
		'    If frmDesigner.PLC1.DescriptionString(Index) = "StationFunction" Then
		'        Load frmNotes
		'        frmNotes.Tag = "crewstation"
		'        frmNotes.Show vbModal, frmDesigner
		'        Set frmNotes = Nothing
		'
		
		'
	End Sub
	Public Sub Properties_Show(ByVal hNode As Integer)
		Dim frmDesigner As Object
		Dim Vehicles As Object
		If hNode <= 0 Then Exit Sub
		
		Dim oNode As Vehicles.cINode
		Dim oDisplay As _cIDisplay
		'UPGRADE_ISSUE: Vehicles.cPropertyItem object was not upgraded. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
		Dim oPropItem As Vehicles.cPropertyItem
		Dim vValue As Object
		Dim lngInterfaceID As Integer
		Const LNG_LENGTH As Short = 4
		Dim index As Integer
		On Error GoTo errDefault
		
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With frmDesigner.PLC1
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.Clear()
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.ShowDescription = True
			'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.Tag = hNode 'CRITICAL - needed so that when proplist attributes for a component are changed, the PLC1 code knows which item is being referenced
		End With
		
		'UPGRADE_WARNING: Couldn't resolve default property of object oNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(oNode, hNode, LNG_LENGTH) '<--- every component in the _Tree_ MUST implement cINode because thats what this pointer is for (if its not going to be rendered in the tree, it doesnt need that interface
		oDisplay = oNode '<-- every component must also obviously implement cIDisplay
		
		Dim oComponent As _cIComponent
		Dim oArmor As cArmor
		Dim oContainer As _cIContainer
		Dim oVehicle As cVehicle
		Dim oBuild As _cIBuild
		Dim oSurface As cSurface
		If Not oDisplay Is Nothing Then
			oPropItem = oDisplay.getfirstpropertyitem
			Do While Not oPropItem Is Nothing
				'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.Caption. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If Not oPropItem.Caption = PROPERTY_HEADER Then
					On Error GoTo errVarName 'todo: this helps us get past bugs while under development, where properties dont exist so callbyname fails
					'NOTE: since the index location of the property in the PLC1  MUST correspond to the  index of the property in the array in our oNode
					' we should fill in a blank line for any property that fails to properly load.
					
					'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.interfaceid. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					lngInterfaceID = oPropItem.interfaceid
					'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.CallByName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Debug.Print("modProperties:Properties_Show() -- InterfaceID = " & lngInterfaceID & " PropertyName = " & oPropItem.CallByName)
					Select Case lngInterfaceID
						'todo: this entire select case needs to be moved to  a seperate function
						'todo: and wouldnt it just be better to call oNode.ClassName and then do the select case by TypeName??
						'      actually, I dont think I can since with composite objects (like armor inside a component) there will exist
						'      multiple interface ID's.  Using typename will not allow us to switch between interfaces.
						'      The real question is, is there a way to get rid of having a huge select case statement?
						Case INTERFACE_NODE
							'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.CallByName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object CallByName(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							vValue = CallByName(oNode, oPropItem.CallByName, CallType.Get)
							
						Case INTERFACE_COMPONENT
							' NOTE: keep this twoard top of select case since its an often used case
							oComponent = oNode
							'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.CallByName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object CallByName(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							vValue = CallByName(oComponent, oPropItem.CallByName, CallType.Get)
							
						Case INTERFACE_ARMOR
							oArmor = oNode
							'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.CallByName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object CallByName(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							vValue = CallByName(oArmor, oPropItem.CallByName, CallType.Get)
							
						Case INTERFACE_CONTAINER
							oContainer = oNode
							'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.CallByName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object CallByName(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							vValue = CallByName(oContainer, oPropItem.CallByName, CallType.Get)
						Case INTERFACE_VEHICLE_DESCRIPTION
							oVehicle = oNode
							'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.CallByName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object CallByName(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							vValue = CallByName(oVehicle.Description, oPropItem.CallByName, CallType.Get)
							
						Case INTERFACE_VEHICLE_VERSION
							' Dim oVehicle As cVehicle
							oVehicle = oNode
							'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.CallByName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object CallByName(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							vValue = CallByName(oVehicle.version, oPropItem.CallByName, CallType.Get)
							
						Case INTERFACE_VEHICLE_AUTHOR
							' Dim oVehicle As cVehicle
							oVehicle = oNode
							'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.CallByName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object CallByName(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							vValue = CallByName(oVehicle.author, oPropItem.CallByName, CallType.Get)
							
						Case INTERFACE_BUILD
							oBuild = oNode
							'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.CallBytype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If oPropItem.CallBytype = 1 Then
								'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.Subscript. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.CallByName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object CallByName(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								vValue = CallByName(oBuild, "get" & oPropItem.CallByName, CallType.Method, oPropItem.Subscript)
								' if its user input, we display the value
								'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								If oPropItem.Datatype = wdDouble Then
									'vValue = vValue
									' else its an option and we use the returned index value to find the string represenation for the selection
									'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								ElseIf oPropItem.Datatype = wdList Then 
									'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.ListItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
									vValue = oPropItem.ListItem(vValue)
								Else
									MsgBox("modProperties.Properties_Show() -- Error: Undefined property type.")
								End If
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.CallByName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object CallByName(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								vValue = CallByName(oBuild, oPropItem.CallByName, CallType.Get)
							End If
							
						Case INTERFACE_SURFACE
							oSurface = oNode
							
							'todo: im potentially going to wind up with the same style If/else block for every interface that has a wdList
							'      is there another way to design this?  Well, maybe its not too many interfaces afterall?  We will see
							'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.Datatype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If oPropItem.Datatype = wdList Then
								'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.CallByName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object CallByName(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								vValue = CallByName(oSurface, oPropItem.CallByName, CallType.Get)
								'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.ListItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								vValue = oPropItem.ListItem(vValue)
							Else
								'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.CallByName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object CallByName(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								'UPGRADE_WARNING: Couldn't resolve default property of object vValue. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
								vValue = CallByName(oSurface, oPropItem.CallByName, CallType.Get)
							End If
							
						Case Else
							'problem
							'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.Caption. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							modHelper.InfoPrint(1, "modProperties:Properties_Show() -- Unsupported Class Interface ID '" & lngInterfaceID & "'  Cannot list property '" & oPropItem.Caption & "'")
					End Select
				End If
				
				On Error GoTo errDefault
				'NOTE: this property must get added here regardless of whether there was a problem accessing its value
				AddPCLproperty(vValue, oPropItem)
				' get the next one
				oPropItem = oDisplay.getnextpropertyitem
			Loop 
			'UPGRADE_WARNING: Couldn't resolve default property of object oNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(oNode, 0, LNG_LENGTH)
			'UPGRADE_NOTE: Object oPropItem may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oPropItem = Nothing
			'UPGRADE_NOTE: Object oDisplay may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			oDisplay = Nothing
		End If
		
		'set the column width of the proplist to always be half the total width
		'todo: this should be done on event when the width of this control changes
		'UPGRADE_WARNING: Couldn't resolve default property of object frmDesigner.PLC1. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		frmDesigner.PLC1.ColumnWidth = frmDesigner.PLC1.Width / (2 * VB6.TwipsPerPixelX)
		Exit Sub
		
errVarName: 
		'UPGRADE_WARNING: Couldn't resolve default property of object oPropItem.CallByName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Debug.Print("modProperties.Properties_Show() -- Could not get value for Variable Name '" & oPropItem.CallByName & "'")
		Resume Next
errDefault: 
		Debug.Print("modProperties:Properties_Show -- Error #" & Err.Number & " " & Err.Description)
		If Not oNode Is Nothing Then
			'UPGRADE_WARNING: Couldn't resolve default property of object oNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(oNode, 0, LNG_LENGTH)
		End If
		'UPGRADE_NOTE: Object oPropItem may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oPropItem = Nothing
		'UPGRADE_NOTE: Object oDisplay may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oDisplay = Nothing
	End Sub
	
	Public Sub DisplayPrintOutput()
		
		Dim sKey As String
		Dim lngDatatype As Integer
		Dim sParentKey As String
		'todo: This could should actually be integrated with the Show_Properties since it has the same
		'      task of determine what type of node we're dealing with
		sKey = p_ActiveNode.Key
		lngDatatype = p_ActiveNode.Datatype
		sParentKey = p_ActiveNode.Parent
		'display the print output in the status bar for this node
		Select Case lngDatatype
			Case Body
				'frmDesigner.StatusBar1.Panels(1).text = m_oCurrentVeh.Body.PrintOutput
			Case VEHICLE_NODE
				' do nothing
				' is a component
			Case PERFORMANCE_NODE
			Case CREW_NODE
			Case POWERSYSTEMS_NODE
				If sParentKey = POWERSYSTEMS_KEY Then
					'frmDesigner.StatusBar1.Panels(1).Text = m_oCurrentVeh.Profiles(sKey).Description
				End If
			Case WEAPON_LINKS_NODE
				If sParentKey = WEAPON_LINKS_KEY Then
					'frmDesigner.StatusBar1.Panels(1).Text = m_oCurrentVeh.WeaponProfiles(sKey).Description & "  --  " & Format(m_oCurrentVeh.WeaponProfiles(sKey).Cost, vbCurrency)
				End If
			Case FUELSYSTEMS_NODE
				If sParentKey = FUELSYSTEMS_KEY Then
					'frmDesigner.StatusBar1.Panels(1).Text = m_oCurrentVeh.Profiles(sKey).Description
				End If
			Case PERFORMANCEAIR To PERFORMANCESPACE
				'frmDesigner.StatusBar1.Panels(1).Text = m_oCurrentVeh.PerformanceProfiles(sKey).Description
				
			Case Else
				' regular components
				'todo: uncomment
				' frmDesigner.StatusBar1.Panels(1).text = m_oCurrentVeh.Components(sKey).PrintOutput
		End Select
	End Sub
	
	Private Sub ShowPropsForPerformanceProfile()
		Dim m_oCurrentVeh As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.ActiveCheckListType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		PopulateCheckList(m_oCurrentVeh.ActiveCheckListType)
		Dim sKey As String
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.ActiveCheckList. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sKey = m_oCurrentVeh.ActiveCheckList
		If sKey <> "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.ActiveCheckListType. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If m_oCurrentVeh.ActiveCheckListType = WEAPON_CHECKLIST Then
				ShowPropsForWeaponLink(sKey)
			Else
				ShowPropsForPerformance(sKey)
			End If
		End If
		
	End Sub
	Private Sub ShowPropsForPowerProfile()
		Dim m_oCurrentVeh As Object
		Dim sKey As String
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.ActiveProfile. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sKey = m_oCurrentVeh.ActiveProfile
		If sKey <> "" Then
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Profiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			m_oCurrentVeh.Profiles(sKey).Show()
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.ActiveProfiletype. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If m_oCurrentVeh.ActiveProfiletype = FUEL_PROFILE Then
				Call ShowLinks(FUEL_PROFILE)
			Else
				Call ShowLinks(POWER_PROFILE)
			End If
		End If
	End Sub
	Private Sub ShowPropsForDescription()
		'these settings are available off of vehicle node
	End Sub
	
	Private Sub ShowPropsForOptions()
		'dont need this... this shows up for vehicle node
	End Sub
	
	Private Sub ShowPropsForCrew()
		Dim m_oCurrentVeh As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With m_oCurrentVeh.crew
			AddPCLproperty("Settings", "", wdText, PROPERTY_HEADER)
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Use Recommended Crew", .UseRecommendedCrew, wdBool, "UseRecommendedCrew")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Occupancy", .Occupancy, wdList, "Occupancy", "short", "long")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Number of Shifts", .numshifts, wdList, "NumShifts", 1, 2, 3, 4, 5, 6, 7, 8, 9)
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Military Vehicle", .MilitaryVehicle, wdBool, "MilitaryVehicle")
			
			AddPCLproperty("Crew Quantities", "", wdText, PROPERTY_HEADER)
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Captains", .numcaptains, wdNumber, "NumCaptains")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Officers", .NumOfficers, wdNumber, "NumOfficers")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Crew Station Operators", .NumCrewStationOperators, wdNumber, "NumCrewStationOperators")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Weapon Loaders", .NumWeaponLoaders, wdNumber, "NumWeaponLoaders")
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Rowers", .NumRowers, wdNumber, "NumRowers")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Sailors", .NumSailors, wdNumber, "NumSailors")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Riggers", .NumRiggers, wdNumber, "NumRiggers")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Fuel Stokers", .NumFuelStokers, wdNumber, "NumFuelStokers")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Mechanics", .NumMechanics, wdNumber, "NumMechanics")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Service Crewmen", .NumServiceCrewmen, wdNumber, "NumServiceCrewmen")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Medics", .NumMedics, wdNumber, "NumMedics")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Scientists", .NumScientists, wdNumber, "NumScientists")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Auxiliary Vehicle Crew", .NumAuxiliaryVehicleCrew, wdNumber, "NumAuxiliaryVehicleCrew")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Stewards", .NumStewards, wdNumber, "NumStewards")
			
			AddPCLproperty("Passenger Quantities", "", wdText, PROPERTY_HEADER)
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Luxury", .NumLuxury, wdNumber, "NumLuxury")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("First Class", .NumFirstClass, wdNumber, "NumFirstClass")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Second Class", .NumSecondClass, wdNumber, "NumSecondClass")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Steerage", .NumSteerage, wdNumber, "NumSteerage")
			
			AddPCLproperty("Stats", "", wdText, PROPERTY_HEADER)
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.crew. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Total Crew + Passengers", .TotalNumberCrewPassengers, wdNumber, "TotalNumberCrewPassengers")
		End With
	End Sub
	
	Private Sub ShowPropsForSurface()
		
	End Sub
	
	Private Sub ShowPropsForStats()
		Dim m_oCurrentVeh As Object
		
		Dim vCategories() As Object
		Dim vSubCategories() As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With m_oCurrentVeh.Description
			Call LoadCategories(vCategories)
			Call LoadSubCategories("Wheeled", vSubCategories)
			
			On Error Resume Next '<-- todo: this is because when you first create a vehicle, the two vCategories() and vSubCategories lines that follow will raise an error
			'Vehicle Description and Authoring Information
			AddPCLproperty("Vehicle Description", "", wdText, PROPERTY_HEADER)
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Name", .NickName, wdText, "NickName")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Class", .Classname, wdText, "ClassName")
			'todo: Hrm... categories suck.  They were implemented originally with the intent that they could be used
			' to filter submissions to the website.  However, i think a better way to handle this is for the user to
			' select the cat/subcat WHEN they want to upload it.  Primarily because these categories can change
			' on the website, and users wont always have proper categories in GVD
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Category", .Category, wdList, "Category", vCategories) 'todo: this one (see above)
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Sub Category", .subcategory, wdList, "subcategory", vSubCategories) 'todo: and this one (see above)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Description", .VehicleDescription, wdText, "VehicleDescription")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Details", .Details, wdText, "Details")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Vision", .Vision, wdText, "Vision")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Header", .Header, wdText, "Header")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Footer", .Footer, wdText, "Footer")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("VehicleImageFileName", .VehicleImageFileName, wdText, "VehicleImageFileName")
			
			AddPCLproperty("Versioning", "", wdText, PROPERTY_HEADER)
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Auto Increment Version", .blnAutoIncrement, wdBool, "blnAutoIncrement")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Vehicle Version", .version, wdText, "version") 'todo: make sure its read only
			
			
			AddPCLproperty("Author Info", "", wdText, PROPERTY_HEADER)
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Name", .author, wdText, "Author")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Email", .email, wdText, "Email")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Description. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Website", .url, wdText, "Url")
		End With
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With m_oCurrentVeh.Options
			
			AddPCLproperty("Miscellaneous", "", wdText, PROPERTY_HEADER)
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Vehicle Crafstmanship", .Quality, wdList, "Quality", "standard", "cheap", "fine", "very fine")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("RollStabilizers", .RollStabilizers, wdBool, "RollStabilizers")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Convertible", .Convertible, wdList, "Convertible", "none", "hardtop", "ragtop")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("UseHardpointMountedWeights", .UseHardpointMountedWeights, wdBool, "UseHardpointMountedWeights")
			
			AddPCLproperty("Payload Settings", "", wdText, PROPERTY_HEADER)
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Use Default Weights", .RecommendedPayload, wdBool, "RecommendedPayload")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Per Person Weight", .PerPersonWeight, wdNumber, "PerPersonWeight")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Per Cargo Weight", .PerCargoWeight, wdNumber, "PerCargoWeight")
			
			AddPCLproperty("Access Space", "", wdText, PROPERTY_HEADER)
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Use Recommended Modifier?", .RecommendedAccessSpace, wdBool, "RecommendedAccessSpace")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Volume Modifier", .AccessSpaceVolumeMod, wdList, "AccessSpaceVolumeMod", 0, 0.25, 0.5, 0.75, 1, 1.25, 1.5, 1.75, 2)
			
			AddPCLproperty("Attachments", "", wdText, PROPERTY_HEADER)
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Pin", .Pin, wdList, "Pin", "none", "standard", "Explosive")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Ram", .Ram, wdBool, "ram")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Bulldozer", .Bulldozer, wdBool, "bulldozer")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Plow", .Plow, wdBool, "Plow")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Options. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Hitch", .Hitch, wdBool, "hitch")
		End With
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With m_oCurrentVeh.surface
			AddPCLproperty("Hull and Hydro Options", "", wdText, PROPERTY_HEADER)
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Streamlining", .StreamLining, wdList, "Streamlining", "none", "fair", "good", "very good", "superior", "excellent", "radical")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Floatation Hull", .FloatationHull, wdBool, "floatationhull")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Submerisible Hull (TL5)", .Submersible, wdBool, "Submersible")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Hydrodynamic Lines", .HydrodynamicLines, wdList, "Hydrodynamiclines", "none", "mediocre", "average", "fine", "very fine", "submarine")
			'AddPCLproperty "Roll Stabilizers (TL7)", .RollStabilizers, wdBool, "rollstabilizers"
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Waterproof", .WaterProof, wdBool, "waterproof")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Sealed (TL5)", .Sealed, wdBool, "Sealed")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Cata/Tri(maran)", .CataTrimaran, wdList, "catatrimaran", "none", "catamaran", "trimaran")
			
			
			
			AddPCLproperty("Concealment", "", wdText, PROPERTY_HEADER)
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Camouflage", .Camouflage, wdBool, "Camouflage")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Infrared Cloaking (TL7)", .infraredcloaking, wdList, "InfraredCloaking", "none", "basic", "radical")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Emission Cloaking (TL8)", .EmissionCloaking, wdList, "EmissionCloaking", "none", "basic", "radical")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Sound Baffling (TL7)", .SoundBaffling, wdList, "SoundBaffling", "none", "basic", "radical")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Stealth (TL7)", .stealth, wdList, "Stealth", "none", "basic", "radical")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Liquid Crystal Skin (TL8)", .LiquidCrystal, wdBool, "LiquidCrystal")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("PsiShielding (TL8)", .PsiShielding, wdBool, "PsiShielding")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Chameleon", .Chameleon, wdList, "Chameleon", "none", "basic", "instant", "intruder")
			
			AddPCLproperty("Magic Levitation", "", wdText, PROPERTY_HEADER)
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Enabled", .bMagicLevitation, wdBool, "bMagicLevitation")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Energy Cost Per Pound", .MagicLevitationEnergyCostPerPound, wdDouble, "MagicLevitationEnergyCostPerPound")
			
			AddPCLproperty("Antigravity Coating", "", wdText, PROPERTY_HEADER)
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Enabled", .bAntigravityCoating, wdBool, "bAntigravityCoating")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Cost Per Sq ft", .AntigravityCoatingCostPerSquareFoot, wdDouble, "AntigravityCoatingCostPerSquareFoot")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Surface Area Useage", .AntigravityCoatingSurfaceAreaUseage, wdList, "AntigravityCoatingSurfaceAreaUseage", "Body", "Vehicle")
			
			AddPCLproperty("Super Science Coating", "", wdText, PROPERTY_HEADER)
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Enabled", .bSuperScienceCoating, wdBool, "bSuperScienceCoating")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Cost Per Sq ft", .SuperScienceCoatingCostPerSquareFoot, wdDouble, "SuperScienceCoatingCostPerSquareFoot")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.surface. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Surface Area Useage", .SuperScienceCoatingSurfaceAreaUseage, wdList, "SuperScienceCoatingSurfaceAreaUseage", "Body", "Vehicle")
			
		End With
		
		'03/24/02 - We dont need to display stats here... it just makes for too much scrolling
		'With m_oCurrentVeh.Stats
		'    'update displays for Price, Health and SizeModifier ,Volume, Weight and Emptyweigh
		'    AddPCLproperty "Statistics", "", wdText, "Disabled"
		'    AddPCLproperty "Price", "$" & Format(.TotalPrice, "standard"), wdText, "Disabled"
		'
		'
		'    AddPCLproperty "Health", .StructuralHealth & " HT", wdText, "Disabled"
		'    AddPCLproperty "SizeMod", .SizeModifier, wdText, "Disabled"
		'    AddPCLproperty "Volume", Format(.TotalVolume, "standard") & " cu ft", wdText, "Disabled"
		'    AddPCLproperty "Area", Format(.totalsurfacearea, "standard") & " sq ft", wdText, "Disabled"
		'    AddPCLproperty "Empty Wt", Format(.EmptyWeight, "standard") & " lbs", wdText, "Disabled"
		'    AddPCLproperty "Empty Mass", Format(.EmptyWeight / 2000, "standard") & " tons", wdText, "Disabled"
		'    AddPCLproperty "Loaded Wt", Format(.LoadedWeight, "standard") & " lbs", wdText, "Disabled"
		'    AddPCLproperty "Loaded Mass", Format(.LoadedMass, "standard") & " tons", wdText, "Disabled"
		'    AddPCLproperty "+Hardpoint Wt", Format(.HLoadedWeight, "standard") & " lbs", wdText, "Disabled"
		'    AddPCLproperty "+Hardpoint Mass", Format(.HLoadedMass, "standard") & " tons", wdText, "Disabled"
		'    AddPCLproperty "Submerged Wt", Format(.SubmergedWeight, "standard") & " lbs", wdText, "Disabled"
		'    AddPCLproperty "Submerged Mass", Format(.SubmergedMass, "standard") & " tons", wdText, "Disabled"
		'    AddPCLproperty "Power Output", Format(.TotalGeneratedPower, "standard") & " kW", wdText, "Disabled"
		'    AddPCLproperty "Power Consumption", Format(.TotalContinuousPowerConsumption, "standard") & " kW", wdText, "Disabled"
		'    AddPCLproperty "Flotation Rating", .FloatationRating, wdText, "Disabled"
		'End With
		'
		
		
		
		'todo: StatusBar1.Panels(2).Text = "Added Wt: " & Format(.OptionsWeight, "standard") & " lbs"
		'    StatusBar1.Panels(1).Text = "Added Cost: $" & Format(.OptionsCost, "standard")
		'    StatusBar1.Panels(3).Text = "Internal Payload Wt: " & Format(.UsualInternalPayload, "standard") & " lbs"
		
		
	End Sub
	
	Public Sub ShowPropsForPerformance(ByVal Key As String)
		Dim m_oCurrentVeh As Object
		
		'PerformanceType is stored in  m_oCurrentVeh.PerformanceProfiles(Key).Datatype
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With m_oCurrentVeh.PerformanceProfiles(Key)
			
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Select Case .Datatype
				'JAW 2000.06.18
				'Added takeoff/land
				
				Case PERFORMANCELEG
					AddPCLproperty("Legged Performance", "", wdText, PROPERTY_HEADER)
					
					AddPCLproperty("Thrust Options", "", wdText, PROPERTY_HEADER)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("PercentThrust", .percentthrust, wdNumber, "PercentThrust")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("TreatTiltRotorsAsPropellers", .TreatTiltRotorsAsPropellers, wdBool, "TreatTiltRotorsAsPropellers")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("AfterBurnersOn", .AfterBurnersOn, wdBool, "AfterBurnersOn")
					
					
					AddPCLproperty("Streamlining", "", wdText, PROPERTY_HEADER)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("HardPointsOn", .HardPointsOn, wdBool, "HardPointsOn")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("WheelsSkidsExtended", .WheelsSkidsExtended, wdBool, "WheelsSkidsExtended")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("PopTurretsExtended", .PopTurretsExtended, wdBool, "PopTurretsExtended")
					
					AddPCLproperty("Weight Percentages", "", wdText, PROPERTY_HEADER)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("% Crew", .PercentCrewWeight, wdNumber, "PercentCrewWeight")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("% Fuel", .PercentFuelWeight, wdNumber, "PercentFuelWeight")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("% Cargo", .PercentCargoWeight, wdNumber, "PercentCargoWeight")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("% Hardpoints/Bays Load", .PercentHardpointWeight, wdNumber, "PercentHardpointWeight")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("% Provisions", .PercentProvisionWeight, wdNumber, "PercentProvisionWeight")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("% Ammunitions", .PercentAmmunitionWeight, wdNumber, "PercentAmmunitionWeight")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("% PercentAuxVehicleWeight", .PercentAuxVehicleWeight, wdNumber, "PercentAuxVehicleWeight")
					
					AddPCLproperty("Statistics", "", wdText, PROPERTY_HEADER)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Total Drivetrain Power", VB6.Format(.gtotalmotivepower, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gSpeed", VB6.Format(.gTopSpeed, "standard") & " mph", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gOffRd", VB6.Format(.gOffRoad, "standard") & " mph", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gAccel", .gAcceleration & " mph/s", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gDecel", .gDeceleration & " mph/s", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gSR", .gStability, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gMR", .gManeuverability, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gP", VB6.Format(.gPressure, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gPDescr", .gPressureDescription, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Advisory", .Advisory, wdText, "Disabled")
					
				Case PERFORMANCETRACK
					AddPCLproperty("Tracked Performance", "", wdText, PROPERTY_HEADER)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gSpeed", VB6.Format(.gTopSpeed, "standard") & " mph", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gOffRd", VB6.Format(.gOffRoad, "standard") & " mph", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gAccel", .gAcceleration & " mph/s", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gDecel", .gDeceleration & " mph/s", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gSR", .gStability, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gMR", .gManeuverability, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gP", VB6.Format(.gPressure, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gPDescr", .gPressureDescription, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Advisory", .Advisory, wdText, "Disabled")
					
				Case PERFORMANCEWHEEL
					AddPCLproperty("Wheeled Performance", "", wdText, PROPERTY_HEADER)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gSpeed", VB6.Format(.gTopSpeed, "standard") & " mph", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gOffRd", VB6.Format(.gOffRoad, "standard") & " mph", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gAccel", .gAcceleration & " mph/s", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gDecel", .gDeceleration & " mph/s", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gSR", .gStability, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gMR", .gManeuverability, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gP", VB6.Format(.gPressure, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gPDescr", .gPressureDescription, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Advisory", .Advisory, wdText, "Disabled")
					
				Case PERFORMANCEFLEX
					AddPCLproperty("Flexibody Performance", "", wdText, PROPERTY_HEADER)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gSpeed", VB6.Format(.gTopSpeed, "standard") & " mph", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gOffRd", VB6.Format(.gOffRoad, "standard") & " mph", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gAccel", .gAcceleration & " mph/s", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gDecel", .gDeceleration & " mph/s", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gSR", .gStability, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gMR", .gManeuverability, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gP", VB6.Format(.gPressure, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gPDescr", .gPressureDescription, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Advisory", .Advisory, wdText, "Disabled")
					
				Case PERFORMANCESKID
					AddPCLproperty("Skid Performance", "", wdText, PROPERTY_HEADER)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gSpeed", VB6.Format(.gTopSpeed, "standard") & " mph", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gOffRd", VB6.Format(.gOffRoad, "standard") & " mph", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gAccel", .gAcceleration & " mph/s", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gDecel", .gDeceleration & " mph/s", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gSR", .gStability, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gMR", .gManeuverability, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gP", VB6.Format(.gPressure, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("gPDescr", .gPressureDescription, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Advisory", .Advisory, wdText, "Disabled")
					
				Case PERFORMANCEAIR
					AddPCLproperty("Air Performance", "", wdText, PROPERTY_HEADER)
					AddPCLproperty("Thrust Options", "", wdText, PROPERTY_HEADER)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("PercentThrust", .percentthrust, wdNumber, "PercentThrust")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("TreatTiltRotorsAsPropellers", .TreatTiltRotorsAsPropellers, wdBool, "TreatTiltRotorsAsPropellers")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("AfterBurnersOn", .AfterBurnersOn, wdBool, "AfterBurnersOn")
					
					
					AddPCLproperty("Streamlining", "", wdText, PROPERTY_HEADER)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("HardPointsOn", .HardPointsOn, wdBool, "HardPointsOn")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("WheelsSkidsExtended", .WheelsSkidsExtended, wdBool, "WheelsSkidsExtended")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("PopTurretsExtended", .PopTurretsExtended, wdBool, "PopTurretsExtended")
					
					AddPCLproperty("Weight Percentages", "", wdText, PROPERTY_HEADER)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("% Crew", .PercentCrewWeight, wdNumber, "PercentCrewWeight")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("% Fuel", .PercentFuelWeight, wdNumber, "PercentFuelWeight")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("% Cargo", .PercentCargoWeight, wdNumber, "PercentCargoWeight")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("% Hardpoints/Bays Load", .PercentHardpointWeight, wdNumber, "PercentHardpointWeight")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("% Provisions", .PercentProvisionWeight, wdNumber, "PercentProvisionWeight")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("% Ammunitions", .PercentAmmunitionWeight, wdNumber, "PercentAmmunitionWeight")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("% PercentAuxVehicleWeight", .PercentAuxVehicleWeight, wdNumber, "PercentAuxVehicleWeight")
					
					AddPCLproperty("Statistics", "", wdText, PROPERTY_HEADER)
					
					'frmDesigner.lblperformance(0).Caption =  "Can Fly?" & vbTab & .aCanFly
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Thrust", VB6.Format(.aMotiveThrust, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Static Lift", VB6.Format(.staticlift, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Drag", .aDrag, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Speed", VB6.Format(.aTopSpeed, "standard") & " mph", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Stall Speed", VB6.Format(.aStallSpeed, "standard") & " mph", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("aAccel", VB6.Format(.aAcceleration, "standard") & " mph/s", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("aDecel", VB6.Format(.aDeceleration, "standard") & " mph/s", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("aMR", .aManeuverability, wdText, "Disabled")
					
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("aSR", .aStability, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("TakeOff Run (yrds)", .aTakeOffRun, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Landing Run (yrds)", .aLandingRun, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Advisory", .Advisory, wdText, "Disabled")
					
					
				Case PERFORMANCEHOVER
					AddPCLproperty("Hovercraft Performance", "", wdText, PROPERTY_HEADER)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hover Alt", .hHoverAltitude & " feet", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Thrust", VB6.Format(.hMotiveThrust, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Static Lift", VB6.Format(.staticlift, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Speed", VB6.Format(.hTopSpeed, "standard") & " mph", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Drag", .hDrag, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("hAccel", VB6.Format(.hAcceleration, "standard") & " mph/s", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("hDecel", VB6.Format(.hDeceleration, "standard") & " mph/s", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("hSR", .hstability, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("hMR", .hmaneuverability & " g", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Advisory", .Advisory, wdText, "Disabled")
					
					
				Case PERFORMANCEMAGLEV
					AddPCLproperty("Mag-Lev Performance", "", wdText, PROPERTY_HEADER)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("mThrust", VB6.Format(.mlMotiveThrust, "standard") & " lbs", wdText, "Disabled")
					'todo: StaticLift?  'AddPCLproperty "Static Lift", Format(.mlstaticlift, "standard") & " lbs", wdText, "Disabled"
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("mSpeed", VB6.Format(.mlTopSpeed, "standard") & " mph", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Stall Speed", VB6.Format(.mlStallSpeed, "standard") & " mph", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("mDrag", .mlDrag, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("mAccel", VB6.Format(.mlAcceleration, "standard") & " mph/s", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("mDecel", VB6.Format(.mlDeceleration, "standard") & " mph/s", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("mSR", .mlStability, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("mMR", .mlManeuverability, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Advisory", .Advisory, wdText, "Disabled")
					
					
				Case PERFORMANCEWATER
					AddPCLproperty("Water Performance", "", wdText, PROPERTY_HEADER)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("wThrust", VB6.Format(.wTotalAquaticThrust, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("wDrag", VB6.Format(.wHydroDrag, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("wSpeed", VB6.Format(.wTopSpeed, "standard") & " mph", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hydro Speed", VB6.Format(.wHydrofoilSpeed, "standard") & " mph", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Planing Speed", VB6.Format(.wPlaningSpeed, "standard") & " mph", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("wAccel", VB6.Format(.wAcceleration, "standard") & " mph/s", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("wDecel", VB6.Format(.wDeceleration, "standard") & " mph/s", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Incr Decel", VB6.Format(.wIDeceleration, "standard") & " mph/s", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("wSR", .wStability & "  " & "wMR: " & .wManeuverability, wdText, "Disabled")
					'AddPCLproperty  "wMR",  .wManeuverability, wdText, "Disabled"
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("wDraft", VB6.Format(.wDraft, "standard") & " feet", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Advisory", .Advisory, wdText, "Disabled")
					
					
				Case PERFORMANCESUB
					AddPCLproperty("Submerged Performance", "", wdText, PROPERTY_HEADER)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("suThrust", VB6.Format(.suTotalAquaticThrust, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("suDrag", .suHydroDrag, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("suSpeed", VB6.Format(.suTopSpeed, "standard") & " mph", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("suAccel", VB6.Format(.suAcceleration, "standard") & " mph/s", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("suDecel", VB6.Format(.suDeceleration, "standard") & " mph/s", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Incr Decel", VB6.Format(.suIDeceleration, "standard") & " mph/s", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("suSR", .suStability, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("suMR", .suManeuverability, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Draft", VB6.Format(.suDraft, "standard") & " feet", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .suCrushDepth = -1 Then
						AddPCLproperty("Crush Depth", "No Crush Depth", wdText, "Disabled")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Crush Depth", VB6.Format(.suCrushDepth, "standard") & " yards", wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Advisory", .Advisory, wdText, "Disabled")
					
					
				Case PERFORMANCESPACE
					AddPCLproperty("Space Performance", "", wdText, PROPERTY_HEADER)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Thrust", VB6.Format(.sMotiveThrust, "standard") & " lbs", wdText, "Disabled")
					'todo: make sure accel is displaying at least 4 digits for space craft
					'which have very slow accel but eventually build up to very fast speeds.
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("sAccel", VB6.Format(.sAccelerationG, "###,###,###.####") & " g", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("sAccel", VB6.Format(.sAccelerationMPH, "standard") & " mph/s", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Turn Around", VB6.Format(.sTurnAroundTime, "standard") & " secs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("sMR", VB6.Format(.sManeuverability, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hyper", VB6.Format(.sHyperSpeed, "standard") & " parsecs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warp", VB6.Format(.sWarpSpeed, "standard") & " parsecs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Jump?", .sJumpDriveable, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Teleport?", .sTeleportationDriveable, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.PerformanceProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Advisory", .Advisory, wdText, "Disabled")
			End Select
		End With
		
	End Sub
	
	Private Sub ShowPropsForGroupComponent(ByVal component As Short, ByVal Key As String)
		Dim m_oCurrentVeh As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With m_oCurrentVeh.Components(Key)
			AddPCLproperty("Settings", "", wdText, "Disabled")
		End With
		
	End Sub
	
	Private Sub ShowPropsForSimpleCustom(ByVal component As Short, ByVal Key As String)
		Dim m_oCurrentVeh As Object
		Select Case component
			
			Case SimpleCustom
				'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				With m_oCurrentVeh.Components(Key)
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("User Cost", .UserCost, wdDouble, "UserCost")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("User Weight", .UserWeight, wdDouble, "UserWeight")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("User Volume", .UserVolume, wdDouble, "UserVolume")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", .PowerReqt, wdDouble, "PowerReqt")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
				End With
		End Select
		
	End Sub
	
	Public Sub ShowPropsForWeaponLink(ByRef sKey As String)
		Dim m_oCurrentVeh As Object
		
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.WeaponProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With m_oCurrentVeh.WeaponProfiles(sKey)
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.WeaponProfiles. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
		End With
	End Sub
	Private Sub ShowPropsForWeaponry1(ByVal component As Short, ByVal Key As String)
		Dim m_oCurrentVeh As Object
		Dim listarray() As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With m_oCurrentVeh.Components(Key)
			
			' Fill the window with properties for the correct Collection item
			Select Case component
				
				Case BlueGreenLaser, RainbowLaser, Laser, UVLaser, IRLaser, Disruptor, ChargedParticleBeam, NeutralParticleBeam, Flamer, Screamer, Stunner, ParalysisBeam, XRayLaser, FusionBeam, GravityBeam, AntiparticleBeam, Graser, Disintegrator, Displacer, BeamedPowerTransmitter, MilitaryParalysisBeam
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Mounting", .Mount, wdList, "Mount", "normal", "concealed")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside")
					If (component = BlueGreenLaser) Or (component = RainbowLaser) Or (component = Disintegrator) Or (component = Flamer) Or (component = Laser) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Energy Drill", .EnergyDrill, wdBool, "EnergyDrill")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Beam Output", .BeamOutput, wdDouble, "BeamOutput")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cyclic Rate", .rof, wdList, "rof", "1/14400", "1/7200", "1/4800", "1/3600", "1/2400", "1/1200", "1/600", "1/300", "1/150", "1/60", "1/30", "1/15", "1/8", "1/4", "1/2", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Range", .Range, wdList, "Range", "close", "normal", "long", "very long", "extreme")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Cells", .PowerCellType, wdList, "PowerCellType", "none", "C cells", "rC cell", "D cells", "rD cells", "E cells", "rE cells")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# Power Cells", .PowerCellQuantity, wdNumber, "PowerCellQuantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("FTL Beam?", .FTL, wdBool, "FTL")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Compact", .Compact, wdBool, "Compact")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Reputation for Quality", .Reliable, wdBool, "Reliable")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Malfunction", .Malfunction, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Type Damage", .TypeDamage, wdText, "Disabled")
					'note: for displacers its radius of effect and not damage
					'note: for paralysis beams its HT penalty and not damage
					'note: for stunners its HT penalty also and not damage
					If (component = Stunner) Or (component = ParalysisBeam) Or (component = MilitaryParalysisBeam) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("HT penalty", .Damage, wdText, "Disabled")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Damage", .Damage, wdText, "Disabled")
					End If
					'note: militaryparalysis,paralysis, disintegrators and displacers have no half damage
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("1/2 Damage (yards)", .halfDamage, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Vacuum 1/2 Damage (yards)", .VacuumHalfDamage, wdText, "Disabled")
					
					If component = Stunner Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Max Range at HT6- (yards)", .MaxRange, wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Max Range at HT7+ (yards)", .MaxRange2, wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Vacuum Max Range at HT6- (yards)", .VacuumMaxRange, wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Vacuum Max Range at HT7+ (yards)", .VacuumMaxRange2, wdText, "Disabled")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Max Range (yards)", .MaxRange, wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Vacuum Max Range (yards)", .VacuumMaxRange, wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Accuracy", .Accuracy, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Snap Shot", .SnapShot, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Rate of Fire", .rof, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Shots", .Shots, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case UnGuidedMissile, UnGuidedTorpedo
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Diameter", .Diameter, wdDouble, "Diameter")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillAmmunitionList)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead", .WarHead, wdList, "Warhead", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .SpaceMissile = False Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Speed (yds per sec)", .Speed, wdDouble, "Speed")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("G's", .Speed, wdDouble, "Speed")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Motor Weight", .MotorWeight, wdDouble, "MotorWeight")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Stealth", .stealth, wdBool, "Stealth")
					'this option only available for relevant missile types
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Space Fairing?", .SpaceMissile, wdBool, "SpaceMissile")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Malfunction", .Malfunction, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .BurstRadius <> -1 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Type Damage 1", .TypeDamage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Damage 1", .Damage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .TypeDamage2 <> "none" Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Type Damage 2", .TypeDamage2, wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Damage 2", .Damage2, wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Accuracy", .Accuracy, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Endurance (seconds)", .Endurance, wdText, "Disabled")
					'only unguided missiles have 1/2 damage
					If component = UnGuidedMissile Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("1/2 Damage (yards)", .halfDamage, wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Max Range (yards)", .MaxRange, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Min Range (yards)", .MinRange, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Motor Cost", "$" & VB6.Format(.MotorCost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Cost", "$" & VB6.Format(.WarheadCost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Weight", VB6.Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Payload Cost", "$" & VB6.Format(.PayloadCost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Payload Weight", VB6.Format(.PayloadWeight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Total Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Total Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case GuidedMissile, GuidedTorpedo
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillGuidanceList)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Guidance System", .GuidanceSystem, wdList, "GuidanceSystem", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillTerminalGuidanceList)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Terminal Guidance", .BrilliantGuidanceSystem, wdList, "BrilliantGuidanceSystem", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cheap Guidance System", .CheapGuidance, wdBool, "CheapGuidance")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Compact Guidance System", .Compact, wdBool, "Compact")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Mid-Course Update", .MidCourseUpdate, wdBool, "MidCourseUpdate")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Pop-Up", .PopUp, wdBool, "Popup")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Diameter", .Diameter, wdDouble, "Diameter")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillAmmunitionList)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead", .WarHead, wdList, "Warhead", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Skill Bonus", .SkillBonus, wdNumber, "SkillBonus")
					'this option only available for relevant missile /torp types
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .SpaceMissile = False Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Speed (yds per sec)", .Speed, wdDouble, "Speed")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("G's", .Speed, wdDouble, "Speed")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Motor Weight", .MotorWeight, wdDouble, "MotorWeight")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Stealth", .stealth, wdBool, "Stealth")
					'this option only available for relevant missile types
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Space Fairing?", .SpaceMissile, wdBool, "SpaceMissile")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Malfunction", .Malfunction, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .BurstRadius <> -1 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Type Damage 1", .TypeDamage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Damage 1", .Damage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .TypeDamage2 <> "none" Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Type Damage 2", .TypeDamage2, wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Damage 2", .Damage2, wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Skill", .Skill, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Endurance (seconds)", .Endurance, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Max Range (yards)", .MaxRange, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Min Range (yards)", .MinRange, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Motor Cost", "$" & VB6.Format(.MotorCost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Cost", "$" & VB6.Format(.WarheadCost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Weight", VB6.Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Guidance System Cost", "$" & VB6.Format(.GuidanceCost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Guidance System Weight", VB6.Format(.GuidanceWeight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Payload Cost", "$" & VB6.Format(.PayloadCost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Payload Weight", VB6.Format(.PayloadWeight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Total Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Total Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
				Case IronBomb, SelfDestructSystem
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Diameter", .Diameter, wdDouble, "Diameter")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillAmmunitionList)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead", .WarHead, wdList, "Warhead", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10")
					'note: these dont have a speed do they?
					'AddPCLproperty "Speed (yds per sec)", .Speed, wdDouble, "Speed"
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Stealth", .stealth, wdBool, "Stealth")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Malfunction", .Malfunction, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .BurstRadius <> -1 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Type Damage 1", .TypeDamage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Damage 1", .Damage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .TypeDamage2 <> "none" Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Type Damage 2", .TypeDamage2, wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Damage 2", .Damage2, wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Accuracy", .Accuracy, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("1/2 Damage (yards)", .halfDamage, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Min Range (yards)", .MinRange, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Weight", VB6.Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.PayloadCost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.PayloadWeight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case RetardedBomb
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Diameter", .Diameter, wdDouble, "Diameter")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillAmmunitionList)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead", .WarHead, wdList, "Warhead", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Speed (yds per sec)", .Speed, wdDouble, "Speed")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Stealth", .stealth, wdBool, "Stealth")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Malfunction", .Malfunction, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .BurstRadius <> -1 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Type Damage 1", .TypeDamage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Damage 1", .Damage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .TypeDamage2 <> "none" Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Type Damage 2", .TypeDamage2, wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Damage 2", .Damage2, wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Accuracy", .Accuracy, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("1/2 Damage (yards)", .halfDamage, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Min Range (yards)", .MinRange, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Payload Cost", "$" & VB6.Format(.PayloadCost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Weight", VB6.Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Payload Weight", VB6.Format(.PayloadWeight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case SmartBomb
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillGuidanceList)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Guidance System", .GuidanceSystem, wdList, "GuidanceSystem", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cheap Guidance System", .CheapGuidance, wdBool, "CheapGuidance")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Compact Guidance System", .Compact, wdBool, "Compact")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Diameter", .Diameter, wdDouble, "Diameter")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillAmmunitionList)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead", .WarHead, wdList, "Warhead", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Skill Bonus", .SkillBonus, wdNumber, "SkillBonus")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Speed (yds per sec)", .Speed, wdDouble, "Speed")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Stealth", .stealth, wdBool, "Stealth")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Malfunction", .Malfunction, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .BurstRadius <> -1 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Type Damage 1", .TypeDamage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Damage 1", .Damage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .TypeDamage2 <> "none" Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Type Damage 2", .TypeDamage2, wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Damage 2", .Damage2, wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Skill", .Skill, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Min Range (yards)", .MinRange, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Guidance System Cost", "$" & VB6.Format(.GuidanceCost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Payload Cost", "$" & VB6.Format(.PayloadCost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Weight", VB6.Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Guidance System Weight", VB6.Format(.GuidanceWeight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Payload Weight", VB6.Format(.PayloadWeight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case ContactMine
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Diameter", .Diameter, wdDouble, "Diameter")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillAmmunitionList)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead", .WarHead, wdList, "Warhead", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Malfunction", .Malfunction, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .BurstRadius <> -1 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Type Damage 1", .TypeDamage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Damage 1", .Damage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .TypeDamage2 <> "none" Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Type Damage 2", .TypeDamage2, wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Damage 2", .Damage2, wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Accuracy", .Accuracy, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Weight", VB6.Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
				Case ProximityMine
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillGuidanceList)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Guidance System", .GuidanceSystem, wdList, "GuidanceSystem", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cheap Guidance System", .CheapGuidance, wdBool, "CheapGuidance")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Compact Guidance System", .Compact, wdBool, "Compact")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Diameter", .Diameter, wdDouble, "Diameter")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillAmmunitionList)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead", .WarHead, wdList, "Warhead", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Skill Bonus", .SkillBonus, wdNumber, "SkillBonus")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Stealth", .stealth, wdBool, "Stealth")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Malfunction", .Malfunction, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .BurstRadius <> -1 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Type Damage 1", .TypeDamage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Damage 1", .Damage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .TypeDamage2 <> "none" Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Type Damage 2", .TypeDamage2, wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Damage 2", .Damage2, wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Skill", .Skill, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Guidance System Cost", "$" & VB6.Format(.GuidanceCost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Payload Cost", "$" & VB6.Format(.PayloadCost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Weight", VB6.Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Guidance System Weight", VB6.Format(.GuidanceWeight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Payload Weight", VB6.Format(.PayloadWeight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case PressureTriggerMine
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Diameter", .Diameter, wdDouble, "Diameter")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Detonation Weight", .DetonationWeight, wdDouble, "DetonationWeight")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillAmmunitionList)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead", .WarHead, wdList, "Warhead", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Malfunction", .Malfunction, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .BurstRadius <> -1 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Type Damage 1", .TypeDamage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Damage 1", .Damage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .TypeDamage2 <> "none" Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Type Damage 2", .TypeDamage2, wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Damage 2", .Damage2, wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Weight", VB6.Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case CommandTriggerMine, SmartTriggerMine
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Diameter", .Diameter, wdDouble, "Diameter")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Parachute Mine?", .Parachute, wdBool, "Parachute")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillAmmunitionList)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead", .WarHead, wdList, "Warhead", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Size", .WarheadSize, wdList, "WarheadSize", "small", "modest", "normal", "big", "huge")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Warheads", .BusMissiles, wdList, "BusMissiles", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Malfunction", .Malfunction, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .BurstRadius <> -1 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Type Damage 1", .TypeDamage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Damage 1", .Damage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .TypeDamage2 <> "none" Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Type Damage 2", .TypeDamage2, wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Damage 2", .Damage2, wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warhead Weight", VB6.Format(.WarheadWeight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case WaterCannon, FlameThrower
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Mounting", .Mount, wdList, "Mount", "normal", "concealed")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Style", .Style, wdList, "Style", "light", "medium", "heavy")
					If component = WaterCannon Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Type of Ammo", .Ammunitiontype, wdList, "AmmunitionType", "water", "acid", "foam")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Shots", .Shots, wdNumber, "Shots")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Malfunction", .Malfunction, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Type Damage", .TypeDamage, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Damage", .Damage, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("1/2 Damage (yards)", .halfDamage, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Max Range (yards)", .MaxRange, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Accuracy", .Accuracy, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Snap Shot", .SnapShot, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Rate of Fire", .rof, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Shots", .Shots, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost Per Shot", .CPS, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight Per Shot", .WPS, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Loaders", .Loaders, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
				Case RevolverLauncher, DisposableLauncher, MuzzleloadingLauncher, BreechloadingLauncher, ManualRepeaterLauncher, SlowAutoLoaderLauncher, FastAutoLoaderLauncher, lightAutomaticLauncher, HeavyAutomaticLauncher
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Mounting", .Mount, wdList, "Mount", "normal", "concealed")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Diameter", .Diameter, wdDouble, "Diameter")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Maximum Load (lbs)", .MaxLoad, wdDouble, "MaxLoad")
					Select Case component
						Case RevolverLauncher, DisposableLauncher, MuzzleloadingLauncher, BreechloadingLauncher
							'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							AddPCLproperty("# of Tubes", .Cylinders, wdNumber, "Cylinders")
					End Select
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Snap Shot", .SnapShot, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Rate of Fire", .rof, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Shots", .Shots, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Loaders", .Loaders, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
			End Select
		End With
		
	End Sub
	
	Private Sub ShowPropsForWeaponry2(ByVal component As Short, ByVal Key As String)
		Dim m_oCurrentVeh As Object
		Dim listarray() As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With m_oCurrentVeh.Components(Key)
			
			' Fill the window with properties for the correct Collection item
			Select Case component
				
				Case StoneThrower, BoltThrower
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Mounting", .Mount, wdList, "Mount", "normal", "concealed")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside")
					If component = StoneThrower Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Mechanism", .Mechanism, wdList, "Mechanism", "spring-powered", "torsion-powered", "counterweight")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Mechanism", .Mechanism, wdList, "Mechanism", "spring-powered", "torsion-powered")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Strength", .Strength, wdNumber, "Strength")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Malfunction", .Malfunction, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Type Damage", .TypeDamage, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Damage", .Damage, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("1/2 Damage (yards)", .halfDamage, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Max Range (yards)", .MaxRange, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Min Range (yards)", .MinRange, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Accuracy", .Accuracy, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Snap Shot", .SnapShot, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Rate of Fire", .rof, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Shots", .Shots, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Reqt. Loaders", .Loaders, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost Per Shot", "$" & VB6.Format(.CPS, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight Per Shot", VB6.Format(.WPS, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume Per Shot", VB6.Format(.VPS, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
				Case RepeatingBoltThrower
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Mounting", .Mount, wdList, "Mount", "normal", "concealed")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Mechanism", .Mechanism, wdList, "Mechanism", "spring-powered", "torsion-powered")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Strength", .Strength, wdNumber, "Strength")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Magazine Capacity", .MagazineCapacity, wdNumber, "MagazineCapacity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Malfunction", .Malfunction, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Type Damage", .TypeDamage, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Damage", .Damage, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("1/2 Damage (yards)", .halfDamage, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Max Range (yards)", .MaxRange, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Min Range (yards)", .MinRange, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Accuracy", .Accuracy, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Snap Shot", .SnapShot, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Rate of Fire", .rof, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Shots", .Shots, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Reqt. Loaders", .Loaders, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost Per Shot", "$" & VB6.Format(.CPS, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight Per Shot", VB6.Format(.WPS, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume Per Shot", VB6.Format(.VPS, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
					
				Case MuzzleLoader, BreechLoader
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Mounting", .Mount, wdList, "Mount", "normal", "concealed")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Bore Size", .BoreSize, wdDouble, "BoreSize")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Technology", .technology, wdList, "Technology", "conventional smoothbore", "conventional rifled", "electromag", "gravitic")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillAmmunitionList)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ammunition", .Ammunitiontype, wdList, "AmmunitionType", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Option", .PowerOption, wdList, "PowerOption", "normal", "low-powered", "extra-low-powered")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Recoiless", .Recoiless, wdBool, "Recoiless")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Barrel Length", .Barrel, wdList, "Barrel", "extremely short", "very short", "short", "medium", "long", "very long", "extremely long")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Fixed Barrels", .Cylinders, wdList, "Cylinders", "1", "2", "3", "4", "5", "6", "7")
					'advanced option not available for unconventinal weapons
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (.technology = "conventional smoothbore") Or (.technology = "conventional rifled") Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Advanced Option", .advancedoption, wdList, "AdvancedOption", "none", "plastic-cased ammunition", "caseless", "liquid propellant", "electrothermal")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.advancedoption = "none"
					End If
					
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Reputation for Quality", .Reliable, wdBool, "Reliable")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (component = ElectricGatling) Or (.technology = "electromag") Or (.Electric) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					End If
					If component = MuzzleLoader Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Carriage Required", .Carriage, wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Malfunction", .Malfunction, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .BurstRadius <> -1 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Type Damage 1", .TypeDamage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Damage 1", .Damage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .TypeDamage2 <> "none" Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Type Damage 2", .TypeDamage2, wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Damage 2", .Damage2, wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("1/2 Damage (yards)", .halfDamage, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Max Range (yards)", .MaxRange, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Accuracy", .Accuracy, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Snap Shot", .SnapShot, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Rate of Fire", .sRoF, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Shots", .Shots, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Reqt. Loaders", .Loaders, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost Per Shot", "$" & VB6.Format(.CPS, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight Per Shot", VB6.Format(.WPS, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume Per Shot", VB6.Format(.VPS, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case ManualRepeater
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Mounting", .Mount, wdList, "Mount", "normal", "concealed")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Bore Size", .BoreSize, wdDouble, "BoreSize")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Technology", .technology, wdList, "Technology", "conventional smoothbore", "conventional rifled", "electromag", "gravitic")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillAmmunitionList)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ammunition", .Ammunitiontype, wdList, "AmmunitionType", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Option", .PowerOption, wdList, "PowerOption", "normal", "low-powered", "extra-low-powered")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Recoiless", .Recoiless, wdBool, "Recoiless")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Barrel Length", .Barrel, wdList, "Barrel", "extremely short", "very short", "short", "medium", "long", "very long", "extremely long")
					'advanced option not available for unconventinal weapons
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (.technology = "conventional smoothbore") Or (.technology = "conventional rifled") Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Advanced Option", .advancedoption, wdList, "AdvancedOption", "none", "plastic-cased ammunition", "caseless", "liquid propellant", "electrothermal")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.advancedoption = "none"
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Box Magazine", .BoxMagazine, wdBool, "BoxMagazine")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Reputation for Quality", .Reliable, wdBool, "Reliable")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (component = ElectricGatling) Or (.technology = "electromag") Or (.Electric) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Malfunction", .Malfunction, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .BurstRadius <> -1 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Type Damage 1", .TypeDamage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Damage 1", .Damage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .TypeDamage2 <> "none" Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Type Damage 2", .TypeDamage2, wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Damage 2", .Damage2, wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("1/2 Damage (yards)", .halfDamage, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Max Range (yards)", .MaxRange, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Accuracy", .Accuracy, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Snap Shot", .SnapShot, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Rate of Fire", .sRoF, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Shots", .Shots, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Reqt. Loaders", .Loaders, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost Per Shot", "$" & VB6.Format(.CPS, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight Per Shot", VB6.Format(.WPS, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume Per Shot", VB6.Format(.VPS, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case Revolver, MechanicalGatling
					'NOTE: These allow for user modifieable Rates of Fire
					'power only needs to be displayed for elec.gat.
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Mounting", .Mount, wdList, "Mount", "normal", "concealed")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Bore Size", .BoreSize, wdDouble, "BoreSize")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Technology", .technology, wdList, "Technology", "conventional smoothbore", "conventional rifled", "electromag", "gravitic")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillAmmunitionList)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ammunition", .Ammunitiontype, wdList, "AmmunitionType", listarray)
					If component = MechanicalGatling Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Operator DX + Skill", .DXPlusSkill, wdDouble, "DXPlusSkill")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Option", .PowerOption, wdList, "PowerOption", "normal", "low-powered", "extra-low-powered")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Recoiless", .Recoiless, wdBool, "Recoiless")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Barrel Length", .Barrel, wdList, "Barrel", "extremely short", "very short", "short", "medium", "long", "very long", "extremely long")
					If component = Revolver Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("# of Cylinders", .Cylinders, wdList, "Cylinders", "3", "4", "5", "6", "7")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("# of Barrels", .Cylinders, wdList, "Cylinders", "3", "4", "5", "6", "7")
					End If
					'advanced option not available for unconventinal weapons
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (.technology = "conventional smoothbore") Or (.technology = "conventional rifled") Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Advanced Option", .advancedoption, wdList, "AdvancedOption", "none", "plastic-cased ammunition", "caseless", "liquid propellant", "electrothermal")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.advancedoption = "none"
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Reputation for Quality", .Reliable, wdBool, "Reliable")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (component = ElectricGatling) Or (.technology = "electromag") Or (.Electric) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Malfunction", .Malfunction, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .BurstRadius <> -1 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Type Damage 1", .TypeDamage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Damage 1", .Damage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .TypeDamage2 <> "none" Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Type Damage 2", .TypeDamage2, wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Damage 2", .Damage2, wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("1/2 Damage (yards)", .halfDamage, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Max Range (yards)", .MaxRange, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Accuracy", .Accuracy, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Snap Shot", .SnapShot, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Rate of Fire", .sRoF, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Shots", .Shots, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Reqt. Loaders", .Loaders, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost Per Shot", "$" & VB6.Format(.CPS, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight Per Shot", VB6.Format(.WPS, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume Per Shot", VB6.Format(.VPS, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
				Case ElectricGatling
					'NOTE: This allows for user modifieable Rate of Fire
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Mounting", .Mount, wdList, "Mount", "normal", "concealed")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Bore Size", .BoreSize, wdDouble, "BoreSize")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Technology", .technology, wdList, "Technology", "conventional smoothbore", "conventional rifled", "electromag", "gravitic")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillAmmunitionList)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ammunition", .Ammunitiontype, wdList, "AmmunitionType", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillRoFList)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Rate of Fire", .dRoF, wdList, "dRoF", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Option", .PowerOption, wdList, "PowerOption", "normal", "low-powered", "extra-low-powered")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Recoiless", .Recoiless, wdBool, "Recoiless")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Barrel Length", .Barrel, wdList, "Barrel", "extremely short", "very short", "short", "medium", "long", "very long", "extremely long")
					If component = Revolver Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("# of Cylinders", .Cylinders, wdList, "Cylinders", "3", "4", "5", "6", "7")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("# of Barrels", .Cylinders, wdList, "Cylinders", "3", "4", "5", "6", "7")
					End If
					'advanced option not available for unconventinal weapons
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (.technology = "conventional smoothbore") Or (.technology = "conventional rifled") Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Advanced Option", .advancedoption, wdList, "AdvancedOption", "none", "plastic-cased ammunition", "caseless", "liquid propellant", "electrothermal")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.advancedoption = "none"
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Reputation for Quality", .Reliable, wdBool, "Reliable")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (component = ElectricGatling) Or (.technology = "electromag") Or (.Electric) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Malfunction", .Malfunction, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .BurstRadius <> -1 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Type Damage 1", .TypeDamage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Damage 1", .Damage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .TypeDamage2 <> "none" Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Type Damage 2", .TypeDamage2, wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Damage 2", .Damage2, wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("1/2 Damage (yards)", .halfDamage, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Max Range (yards)", .MaxRange, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Accuracy", .Accuracy, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Snap Shot", .SnapShot, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Shots", .Shots, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Reqt. Loaders", .Loaders, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost Per Shot", "$" & VB6.Format(.CPS, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight Per Shot", VB6.Format(.WPS, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume Per Shot", VB6.Format(.VPS, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case SlowAutoloader, FastAutoloader
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Mounting", .Mount, wdList, "Mount", "normal", "concealed")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Bore Size", .BoreSize, wdDouble, "BoreSize")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Technology", .technology, wdList, "Technology", "conventional smoothbore", "conventional rifled", "electromag", "gravitic")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillAmmunitionList)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ammunition", .Ammunitiontype, wdList, "AmmunitionType", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Option", .PowerOption, wdList, "PowerOption", "normal", "low-powered", "extra-low-powered")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Recoiless", .Recoiless, wdBool, "Recoiless")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Barrel Length", .Barrel, wdList, "Barrel", "extremely short", "very short", "short", "medium", "long", "very long", "extremely long")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Electric Loading", .Electric, wdBool, "Electric")
					'advanced option not available for unconventinal weapons
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (.technology = "conventional smoothbore") Or (.technology = "conventional rifled") Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Advanced Option", .advancedoption, wdList, "AdvancedOption", "none", "plastic-cased ammunition", "caseless", "liquid propellant", "electrothermal")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.advancedoption = "none"
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Reputation for Quality", .Reliable, wdBool, "Reliable")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (component = ElectricGatling) Or (.technology = "electromag") Or (.technology = "gravitic") Or (.Electric) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Malfunction", .Malfunction, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .BurstRadius <> -1 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Type Damage 1", .TypeDamage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Damage 1", .Damage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .TypeDamage2 <> "none" Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Type Damage 2", .TypeDamage2, wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Damage 2", .Damage2, wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("1/2 Damage (yards)", .halfDamage, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Max Range (yards)", .MaxRange, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Accuracy", .Accuracy, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Snap Shot", .SnapShot, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Rate of Fire", .sRoF, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Shots", .Shots, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Reqt. Loaders", .Loaders, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost Per Shot", "$" & VB6.Format(.CPS, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight Per Shot", VB6.Format(.WPS, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume Per Shot", VB6.Format(.VPS, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
				Case lightAutomatic, HeavyAutomatic
					'note: these allow for user edit-able Rates of Fire
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quality", .Quality, wdList, "Quality", "normal", "cheap", "fine (accurate)", "very fine (accurate)", "fine (reliable)")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Mounting", .Mount, wdList, "Mount", "normal", "concealed")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Bore Size", .BoreSize, wdDouble, "BoreSize")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Technology", .technology, wdList, "Technology", "conventional smoothbore", "conventional rifled", "electromag", "gravitic")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillAmmunitionList)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ammunition", .Ammunitiontype, wdList, "AmmunitionType", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillRoFList)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Rate of Fire", .dRoF, wdList, "dRoF", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Option", .PowerOption, wdList, "PowerOption", "normal", "low-powered", "extra-low-powered")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Recoiless", .Recoiless, wdBool, "Recoiless")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Barrel Length", .Barrel, wdList, "Barrel", "extremely short", "very short", "short", "medium", "long", "very long", "extremely long")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Electric Loading", .Electric, wdBool, "Electric")
					'advanced option not available for unconventinal weapons
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (.technology = "conventional smoothbore") Or (.technology = "conventional rifled") Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Advanced Option", .advancedoption, wdList, "AdvancedOption", "none", "plastic-cased ammunition", "caseless", "liquid propellant", "electrothermal")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.advancedoption = "none"
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Reputation for Quality", .Reliable, wdBool, "Reliable")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (component = ElectricGatling) Or (.technology = "electromag") Or (.Electric) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Malfunction", .Malfunction, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .BurstRadius <> -1 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Burst Radius(yds)", .BurstRadius, wdNumber, "BurstRadius")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Type Damage 1", .TypeDamage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Damage 1", .Damage1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .TypeDamage2 <> "none" Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Type Damage 2", .TypeDamage2, wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Damage 2", .Damage2, wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("1/2 Damage (yards)", .halfDamage, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Max Range (yards)", .MaxRange, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Accuracy", .Accuracy, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Snap Shot", .SnapShot, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Shots", .Shots, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Reqt. Loaders", .Loaders, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost Per Shot", "$" & VB6.Format(.CPS, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight Per Shot", VB6.Format(.WPS, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume Per Shot", VB6.Format(.VPS, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
				Case AntiBlastMagazine
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					
				Case UniversalMount, CasemateMount, DoorMount, Cyberslave, FullStabilizationGear, PartialStabilizationGear
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
					
				Case WeaponBay, HardPoint
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Index", .index, wdNumber, "Index")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Maximum Load (lbs)", .loadcapacity, wdDouble, "LoadCapacity")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Abbreviation", .abbrev, wdText, "Abbrev")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					If component = WeaponBay Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					End If
					
					
				Case Ammunition
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Shots", .NumShots, wdNumber, "NumShots")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Lock Ammo Settings", .Locked, wdBool, "Locked")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ammo Type", .Ammunitiontype, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("CPS", "$" & VB6.Format(.CPS, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("WPS", VB6.Format(.WPS, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("VPS", VB6.Format(.VPS, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
			End Select
			
		End With
	End Sub
	
	Private Sub ShowPropsForBody()
		Dim m_oCurrentVeh As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With m_oCurrentVeh.Components(BODY_KEY)
			AddPCLproperty("Settings", "", wdText, "Disabled")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Compartmentalization", .Compartmentalization, wdList, "Compartmentalization", "none", "heavy", "total")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Flexibody Option", .FlexibodyOption, wdBool, "FlexibodyOption")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Improved Flexibody Suspension", .ImprovedSuspension, wdBool, "ImprovedSuspension")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Lifting Body", .liftingbody, wdBool, "LiftingBody")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Top Deck", .TopDeck, wdBool, "Topdeck")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("% Covered Deck", .PercentCovered, wdNumber, "PercentCovered")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("% Flight Deck", .PercentFlightDeck, wdNumber, "PercentFlightDeck")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Flight Deck Option", .flightdeckoption, wdList, "flightdeckoption", "none", "landing pad", "angled flight deck")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Slope Right", .SlopeR, wdList, "sloper", "none", "30 degrees", "60 degrees")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Slope Left", .slopel, wdList, "slopel", "none", "30 degrees", "60 degrees")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Slope Front", .slopef, wdList, "slopeF", "none", "30 degrees", "60 degrees")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Slope Back", .slopeb, wdList, "slopeb", "none", "30 degrees", "60 degrees")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Empty Space", .EmptySpace, wdDouble, "EmptySpace")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Responsive", .Responsive, wdBool, "Responsive")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Robotic", .Robotic, wdBool, "Robotic")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Biomechanical", .Biomechanical, wdBool, "Biomechanical")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Living Metal", .LivingMetal, wdBool, "LivingMetal")
			
			AddPCLproperty("Statistics", "", wdText, "Disabled")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Abbreviation", .abbrev, wdText, "Abbrev")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Top Deck Area", VB6.Format(.TotalDeckArea, "standard") & " sq ft", wdText, "Disabled")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Flight Deck Length", VB6.Format(.flightdecklength, "standard") & " ft", wdText, "Disabled")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Flight Deck Area", VB6.Format(.FlightDeckArea, "standard") & " sq ft", wdText, "Disabled")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Covered Deck Area", VB6.Format(.covereddeckarea, "standard") & " sq ft", wdText, "Disabled")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Deck Cost", "$" & VB6.Format(.DeckCost, "standard"), wdText, "Disabled")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Deck Weight", VB6.Format(.DeckWeight, "standard") & " lbs", wdText, "Disabled")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Body Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Body Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Body Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Body Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Access Space", VB6.Format(.AccessSpace, "standard") & " cf", wdText, "Disabled")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Minimum Volume", .MinimumVolume, wdText, "Disabled")
			'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
		End With
		
	End Sub
	Private Sub ShowPropsForSubAssemblies(ByVal component As Short, ByVal Key As String)
		Dim m_oCurrentVeh As Object
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With m_oCurrentVeh.Components(Key)
			
			' Fill the window with properties for the correct Collection item
			Select Case component
				
				
				Case Wheel
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Wheel Type", .subtype, wdList, "Subtype", "standard", "small", "heavy", "railway", "off-road", "retractable")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Wheels", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Improved Suspension", .ImprovedSuspension, wdBool, "ImprovedSuspension")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Retract Location", .RetractLocation, wdList, "RetractLocation", "none", "body", "body & wings")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Wheel Blades", .Wheelblades, wdList, "Wheelblades", "none", "fixed", "rectractable")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Snow Tires", .snowtires, wdBool, "Snowtires")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Racing Tires", .racingtires, wdBool, "RacingTires")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Puncture Resistant", .PunctureResistant, wdBool, "PunctureResistant")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Improved Brakes", .ImprovedBrakes, wdBool, "ImprovedBrakes")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("All Wheel Steering", .AllwheelSteering, wdBool, "AllWheelSteering")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Smart Wheels", .Smartwheels, wdBool, "SmartWheels")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Empty Space", .EmptySpace, wdDouble, "EmptySpace")
					'note, no empty space allowed
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Responsive", .Responsive, wdBool, "Responsive")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Robotic", .Robotic, wdBool, "Robotic")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Biomechanical", .Biomechanical, wdBool, "Biomechanical")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Living Metal", .LivingMetal, wdBool, "LivingMetal")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Abbreviation", .abbrev, wdText, "Abbrev")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
				Case Skid
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Skids", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Improved Suspension", .ImprovedSuspension, wdBool, "ImprovedSuspension")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Retract Location", .RetractLocation, wdList, "RetractLocation", "none", "body", "body & wings")
					'note, no empty space allowed
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Responsive", .Responsive, wdBool, "Responsive")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Robotic", .Robotic, wdBool, "Robotic")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Biomechanical", .Biomechanical, wdBool, "Biomechanical")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Living Metal", .LivingMetal, wdBool, "LivingMetal")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Abbreviation", .abbrev, wdText, "Abbrev")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case Track
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Track Type", .subtype, wdList, "SubType", "tracks", "halftracks", "skitracks")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Tracks", .Quantity, wdList, "Quantity", 2, 4)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Improved Suspension", .ImprovedSuspension, wdBool, "ImprovedSuspension")
					'note, no empty space allowed
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Responsive", .Responsive, wdBool, "Responsive")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Robotic", .Robotic, wdBool, "Robotic")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Biomechanical", .Biomechanical, wdBool, "Biomechanical")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Living Metal", .LivingMetal, wdBool, "LivingMetal")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Abbreviation", .abbrev, wdText, "Abbrev")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case Arm
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Index", .index, wdNumber, "Index")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Orientation", .Orientation, wdList, "Orientation", "top", "underside", "front", "back", "left", "right")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Empty Space", .EmptySpace, wdDouble, "EmptySpace")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Responsive", .Responsive, wdBool, "Responsive")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Robotic", .Robotic, wdBool, "Robotic")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Biomechanical", .Biomechanical, wdBool, "Biomechanical")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Living Metal", .LivingMetal, wdBool, "LivingMetal")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Abbreviation", .abbrev, wdText, "Abbrev")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Access Space", VB6.Format(.AccessSpace, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case Leg
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Index", .index, wdNumber, "Index")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Improved Suspension", .ImprovedSuspension, wdBool, "ImprovedSuspension")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Empty Space", .EmptySpace, wdDouble, "EmptySpace")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Responsive", .Responsive, wdBool, "Responsive")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Robotic", .Robotic, wdBool, "Robotic")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Biomechanical", .Biomechanical, wdBool, "Biomechanical")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Living Metal", .LivingMetal, wdBool, "LivingMetal")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Abbreviation", .abbrev, wdText, "Abbrev")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Access Space", VB6.Format(.AccessSpace, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case Wing
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Index", .index, wdNumber, "Index")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Orientation", .Orientation, wdList, "Orientation", "left", "right")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Wing Type", .subtype, wdList, "SubType", "standard", "STOL", "biplane", "triplane", "high agility", "flarecraft", "stub")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Controlled Instability", .ControlledInstability, wdBool, "ControlledInstability")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Folding Wings", .Folding, wdBool, "Folding")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Variable Sweep Wings", .VariableSweep, wdList, "VariableSweep", "none", "manual", "automatic")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Empty Space", .EmptySpace, wdDouble, "EmptySpace")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Responsive", .Responsive, wdBool, "Responsive")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Robotic", .Robotic, wdBool, "Robotic")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Biomechanical", .Biomechanical, wdBool, "Biomechanical")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Living Metal", .LivingMetal, wdBool, "LivingMetal")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Abbreviation", .abbrev, wdText, "Abbrev")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Access Space", VB6.Format(.AccessSpace, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case AutogyroRotor, TTRotor, CARotor, MMRotor
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Index", .index, wdNumber, "Index")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Controlled Instability", .ControlledInstability, wdBool, "ControlledInstability")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Folding Rotors", .Folding, wdBool, "Folding")
					'note, no empty space allowed
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Responsive", .Responsive, wdBool, "Responsive")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Robotic", .Robotic, wdBool, "Robotic")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Biomechanical", .Biomechanical, wdBool, "Biomechanical")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Living Metal", .LivingMetal, wdBool, "LivingMetal")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Abbreviation", .abbrev, wdText, "Abbrev")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case Hydrofoil
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Index", .index, wdNumber, "Index")
					'note, no empty space allowed
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Responsive", .Responsive, wdBool, "Responsive")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Robotic", .Robotic, wdBool, "Robotic")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Biomechanical", .Biomechanical, wdBool, "Biomechanical")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Living Metal", .LivingMetal, wdBool, "LivingMetal")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Abbreviation", .abbrev, wdText, "Abbrev")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'access space because Aquatic propulsion can be placed in them
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Access Space", VB6.Format(.AccessSpace, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case Hovercraft
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Index", .index, wdNumber, "Index")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hovercraft Type", .subtype, wdList, "SubType", "GEV skirt", "SEV sidewalls")
					'note, no empty space allowed
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Responsive", .Responsive, wdBool, "Responsive")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Robotic", .Robotic, wdBool, "Robotic")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Biomechanical", .Biomechanical, wdBool, "Biomechanical")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Living Metal", .LivingMetal, wdBool, "LivingMetal")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Abbreviation", .abbrev, wdText, "Abbrev")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case Superstructure
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Index", .index, wdNumber, "Index")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Orientation", .Orientation, wdList, "Orientation", "top", "underside", "front", "back", "left", "right")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Compartmentalization", .Compartmentalization, wdList, "Compartmentalization", "none", "heavy", "total")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Top Deck", .TopDeck, wdBool, "TopDeck")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("% Covered Deck", .PercentCovered, wdNumber, "PercentCovered")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("% Flight Deck", .PercentFlightDeck, wdNumber, "PercentFlightDeck")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Flight Deck Option", .flightdeckoption, wdList, "FlightDeckOption", "none", "landing pad", "angled flight deck")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Slope Right", .SlopeR, wdList, "sloper", "none", "30 degrees", "60 degrees")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Slope Left", .slopel, wdList, "slopel", "none", "30 degrees", "60 degrees")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Slope Front", .slopef, wdList, "slopeF", "none", "30 degrees", "60 degrees")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Slope Back", .slopeb, wdList, "slopeb", "none", "30 degrees", "60 degrees")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Empty Space", .EmptySpace, wdDouble, "EmptySpace")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Responsive", .Responsive, wdBool, "Responsive")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Robotic", .Robotic, wdBool, "Robotic")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Biomechanical", .Biomechanical, wdBool, "Biomechanical")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Living Metal", .LivingMetal, wdBool, "LivingMetal")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Abbreviation", .abbrev, wdText, "Abbrev")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Top Deck Area", VB6.Format(.TotalDeckArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Flight Deck Length", VB6.Format(.flightdecklength, "standard") & " ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Flight Deck Area", VB6.Format(.FlightDeckArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Covered Deck Area", VB6.Format(.covereddeckarea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Deck Cost", "$" & VB6.Format(.DeckCost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Deck Weight", VB6.Format(.DeckWeight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Access Space", VB6.Format(.AccessSpace, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case OpenMount
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Index", .index, wdNumber, "Index")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Orientation", .Orientation, wdList, "Orientation", "top", "underside", "front", "back", "left", "right")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Rotation Type", .Rotation, wdList, "Rotation", "none", "full", "limited")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Empty Space", .EmptySpace, wdDouble, "EmptySpace")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Abbreviation", .abbrev, wdText, "Abbrev")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case Mast
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Index", .index, wdNumber, "Index")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Masts", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Height", .Height, wdNumber, "Height")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Materials", .Materials, wdList, "Materials", "wood", "metal")
					'note no empty space allowed
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Abbreviation", .abbrev, wdText, "Abbrev")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case Pod
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Index", .index, wdNumber, "Index")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Orientation", .Orientation, wdList, "Orientation", "top", "underside", "front", "back", "left", "right")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Empty Space", .EmptySpace, wdDouble, "EmptySpace")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Responsive", .Responsive, wdBool, "Responsive")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Robotic", .Robotic, wdBool, "Robotic")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Biomechanical", .Biomechanical, wdBool, "Biomechanical")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Living Metal", .LivingMetal, wdBool, "LivingMetal")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Abbreviation", .abbrev, wdText, "Abbrev")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Access Space", VB6.Format(.AccessSpace, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case Turret, Popturret
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'note: only turrets and popturrets will have a "rotation space" statistic
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Index", .index, wdNumber, "Index")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Orientation", .Orientation, wdList, "Orientation", "top", "underside", "front", "back", "left", "right")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Rotation", .Rotation, wdList, "Rotation", "none", "limited", "full")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Compartmentalization", .Compartmentalization, wdList, "Compartmentalization", "none", "heavy", "total")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Slope Right", .SlopeR, wdList, "sloper", "none", "30 degrees", "60 degrees")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Slope Left", .slopel, wdList, "slopel", "none", "30 degrees", "60 degrees")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Slope Front", .slopef, wdList, "slopeF", "none", "30 degrees", "60 degrees")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Slope Back", .slopeb, wdList, "slopeb", "none", "30 degrees", "60 degrees")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Empty Space", .EmptySpace, wdDouble, "EmptySpace")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Responsive", .Responsive, wdBool, "Responsive")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Robotic", .Robotic, wdBool, "Robotic")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Biomechanical", .Biomechanical, wdBool, "Biomechanical")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Living Metal", .LivingMetal, wdBool, "LivingMetal")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Abbreviation", .abbrev, wdText, "Abbrev")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Rotation Space", .RotationSpace, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Access Space", VB6.Format(.AccessSpace, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case Gasbag
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Index", .index, wdNumber, "Index")
					'note no empty space allowed
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Abbreviation", .abbrev, wdText, "Abbrev")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case Cargo
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'AddPCLproperty "Index", .Index, wdNumber, "Index"
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cargo Type", .subtype, wdList, "Subtype", "standard", "hidden", "open")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cargo Room", .CargoSpace, wdDouble, "CargoSpace")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Empty Weight", .Weight, wdDouble, "Weight")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight Per cf", .WeightPerCubicFoot, wdDouble, "WeightPerCubicFoot")
					'note no empty space allowed
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Compartment Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cargo Weight", VB6.Format(.CargoWeight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					
					
				Case equipmentPod
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Index", .index, wdNumber, "Index")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Empty Space", .EmptySpace, wdDouble, "EmptySpace")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Responsive", .Responsive, wdBool, "Responsive")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Robotic", .Robotic, wdBool, "Robotic")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Biomechanical", .Biomechanical, wdBool, "Biomechanical")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Living Metal", .LivingMetal, wdBool, "LivingMetal")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Abbreviation", .abbrev, wdText, "Abbrev")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Access Space", VB6.Format(.AccessSpace, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case SideCar
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Index", .index, wdNumber, "Index")
					'note, no empty space allowed
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Responsive", .Responsive, wdBool, "Responsive")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Robotic", .Robotic, wdBool, "Robotic")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Biomechanical", .Biomechanical, wdBool, "Biomechanical")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Living Metal", .LivingMetal, wdBool, "LivingMetal")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Abbreviation", .abbrev, wdText, "Abbrev")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Access Space", VB6.Format(.AccessSpace, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case SolarPanel
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'AddPCLproperty "Index", .Index, wdNumber, "Index"
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", .SurfaceArea, wdDouble, "SurfaceArea")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Retractable?", .Retractable, wdBool, "Retractable")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Frame Strength", .FrameStrength, wdList, "FrameStrength", "super-light", "extra-light", "light", "medium", "heavy", "extra-heavy")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Materials", .Materials, wdList, "Materials", "very cheap", "cheap", "standard", "expensive", "very expensive", "advanced")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Responsive", .Responsive, wdBool, "Responsive")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Robotic", .Robotic, wdBool, "Robotic")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Biomechanical", .Biomechanical, wdBool, "Biomechanical")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Living Metal", .LivingMetal, wdBool, "LivingMetal")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'AddPCLproperty "Abbreviation", .abbrev, wdText, "Abbrev"
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case SolarCellArray
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Percent Area Covered", .PercentCovered, wdNumber, "PercentCovered")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Output", VB6.Format(.Output, "standard") & " kw", wdText, "Disabled")
					'AddPCLproperty "Endurance", .Endurance & " yrs", wdText, "Disabled"
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Consumed", VB6.Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Remaining", VB6.Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled")
			End Select
		End With
	End Sub
	
	Private Sub ShowPropsForPropulsion(ByVal component As Short, ByVal Key As String)
		Dim m_oCurrentVeh As Object
		'////////////////////////////////////////////
		'Propulsion Systems
		'////////////////////////////////////////////
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With m_oCurrentVeh.Components(Key)
			
			' Fill the window with properties for the correct Collection item
			Select Case component
				Case WheeledDrivetrain, AllWheelDriveWheeledDrivetrain, TrackedDrivetrain, LegDrivetrain, FlexibodyDrivetrain
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					If component = LegDrivetrain Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Motive Power (per motor)", .motivepower, wdDouble, "MotivePower")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Motive Power", .motivepower, wdDouble, "MotivePower")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					If component = LegDrivetrain Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Volume per leg:", VB6.Format(.Volume, "standard"), wdText, "Disabled")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case OrnithopterDrivetrain, TTRRotorDrivetrain, CARRotorDrivetrain
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					If component = OrnithopterDrivetrain Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Motive Power (per motor)", .motivepower, wdDouble, "MotivePower")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Motive Power", .motivepower, wdDouble, "MotivePower")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Lift", VB6.Format(.Lift, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Motive Thrust", VB6.Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case MMRRotorDrivetrain
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tilt Rotor?", .TiltRotor, wdBool, "TiltRotor")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Motive Power", .motivepower, wdDouble, "MotivePower")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Lift", VB6.Format(.Lift, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Motive Thrust", VB6.Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case AerialPropeller
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Motive Power", .motivepower, wdDouble, "MotivePower")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Motive Thrust", VB6.Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
					'AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
					'AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
					
					
				Case DuctedFan
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Motive Power", .motivepower, wdDouble, "MotivePower")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hover Fan", .HoverFan, wdBool, "HoverFan")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Lift Engine", .LiftEngine, wdBool, "LiftEngine")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Vectored Thrust", .VectoredThrust, wdBool, "VectoredThrust")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Motive Thrust", VB6.Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
					
				Case PaddleWheel, ScrewPropeller, lightScrewPropeller, DuctedPropeller, Hydrojet, MHDTunnel
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Motive Power", .motivepower, wdDouble, "MotivePower")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Aquatic Motive Thrust", VB6.Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case RopeHarness, YokeandPoleHarness, ShaftandCollarHarness, WhiffletreeHarness
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Animal Type", .subtype, wdList, "SubType", "Land Animal", "Swimming Animal", "Flying Animal")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Animal Description", .AnimalDescription, wdText, "AnimalDescription")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Strength per Animal", .BeastST, wdNumber, "BeastST")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hexes Per Animal", .Hexes, wdNumber, "Hexes")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Number of Animals", .Quantity, wdNumber, "Quantity")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .subtype = "Land Animal" Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Motive Power", .motivepower, wdText, "Disabled")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Motive Thrust", VB6.Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Total Hexes of Animals", .TotalHexes, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Move per Animal", .Move, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Speed per Animal", .Speed, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					
					
					
				Case RowingPositions
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Positions", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Avg. ST per Position", .RowerST, wdNumber, "RowerST")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR per Position", .dr, wdNumber, "DR")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Motive Thrust", VB6.Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case FullRig, SquareRig, ForeandAftRig, AerialSail, AerialSailForeAftRig
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Sail Material", .material, wdList, "Material", "cloth", "synthetic", "bioplas")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Wind", .Wind, wdList, "Wind", "calm", "light air", "light breeze", "gentle breeze", "moderate breeze", "fresh breeze", "strong breeze", "gale force winds")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Motive Thrust", VB6.Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case lightSail
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Sail Size (sq mi)", .SurfaceArea, wdDouble, "SurfaceArea")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("AU Distance", .AUDistance, wdDouble, "AUDistance")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Motive Thrust", VB6.Format(.Thrust, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case Turbojet, Turbofan, Ramjet, TurboRamjet, Hyperfan
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Thrust (lbs)", .DesiredThrust, wdDouble, "DesiredThrust")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Afterburner", .Afterburner, wdBool, "Afterburner")
					If component <> Ramjet Then '//ramjets cant be lift engines because they need air travelling through them in forward motion
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Lift Engine", .LiftEngine, wdBool, "LiftEngine")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Vectored Thrust", .VectoredThrust, wdBool, "VectoredThrust")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Motive Thrust", VB6.Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Output", VB6.Format(.Output, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Consumed", VB6.Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Remaining", VB6.Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Fuel Consumption", VB6.Format(.FuelConsumption, "standard") & " gph", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("AB Thrust", VB6.Format(.ABThrust, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("AB Lit Fuel Consumption", VB6.Format(.ABConsumption, "standard") & " gph", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case FusionAirRam 'only jet engine that cant use afterburner
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Thrust (lbs)", .DesiredThrust, wdDouble, "DesiredThrust")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Lift Engine", .LiftEngine, wdBool, "LiftEngine")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Vectored Thrust", .VectoredThrust, wdBool, "VectoredThrust")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Output", VB6.Format(.Output, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Consumed", VB6.Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Remaining", VB6.Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled")
					
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Motive Thrust", VB6.Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Endurance", VB6.Format(.FuelConsumption, "standard") & " yrs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case StandardThruster, SuperThruster, MegaThruster
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Thrust (lbs)", .DesiredThrust, wdDouble, "DesiredThrust")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Lift Engine", .LiftEngine, wdBool, "LiftEngine")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Vectored Thrust", .VectoredThrust, wdBool, "VectoredThrust")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Motive Thrust", VB6.Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case LiquidFuelRocket, MOXRocket, IonDrive, FissionRocket, FusionRocket, OptimizedFusion, AntimatterThermal
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Thrust (lbs)", .DesiredThrust, wdDouble, "DesiredThrust")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Lift Engine", .LiftEngine, wdBool, "LiftEngine")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Vectored Thrust", .VectoredThrust, wdBool, "VectoredThrust")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Motive Thrust", VB6.Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Fuel Consumption", VB6.Format(.FuelConsumption, "standard") & " gph", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case AntimatterPion
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Thrust (lbs)", .DesiredThrust, wdDouble, "DesiredThrust")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Lift Engine", .LiftEngine, wdBool, "LiftEngine")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Vectored Thrust", .VectoredThrust, wdBool, "VectoredThrust")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Motive Thrust", VB6.Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Antimatter Fuel Consumption", VB6.Format(.FuelConsumption, "standard") & " grams per hour", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hydrogen Fuel Consumption", VB6.Format(.FuelConsumption2, "standard") & " gph", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case SolidRocketEngine
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Thrust (lbs)", .DesiredThrust, wdDouble, "DesiredThrust")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Burn Time (mins)", .BurnTime, wdDouble, "BurnTime")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Lift Engine", .LiftEngine, wdBool, "LiftEngine")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Vectored Thrust", .VectoredThrust, wdBool, "VectoredThrust")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Motive Thrust", VB6.Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case OrionEngine
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Pulse Rate (bps)", .PulseRate, wdDouble, "PulseRate")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Bomb Size (kt)", .BombSize, wdDouble, "BombSize")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Bombs", .NumBombs, wdNumber, "NumBombs")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Lift Engine", .LiftEngine, wdBool, "LiftEngine")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Motive Thrust", VB6.Format(.MotiveThrust, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Thrust Time (secs)", .ThrustTime, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Bomb Weight", .BombWeight, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Bomb Cost", .BombCost, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Bomb Volume", .BombVolume, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Total Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Total Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Total Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case MagLevLifter
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Lift (lbs)", .DesiredLift, wdDouble, "DesiredLift")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Total lift", VB6.Format(.Lift, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case JumpDrive, TeleportationDrive
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Capacity (tons)", .DesiredCapacity, wdDouble, "DesiredCapacity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Total Capacity", VB6.Format(.capacity, "standard") & " tons", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case Hyperdrive
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Capacity (tons)", .DesiredCapacity, wdDouble, "DesiredCapacity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Total Capacity", VB6.Format(.capacity, "standard") & " tons", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Initial Power", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Sustained Power", VB6.Format(.SustainedPower, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case WarpDrive
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Warp Thrust Factor", .DesiredCapacity, wdDouble, "DesiredCapacity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Total WTF", VB6.Format(.capacity, "standard") & " WTF", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case SubQuantumConveyor, QuantumConveyor, TwoQuantumConveyor
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Max Transport Weight (lbs)", .DesiredCapacity, wdDouble, "desiredCapacity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Total Transport Weight", VB6.Format(.capacity, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
					'/////////////////////////////////////////
					' Aerostatic Lift Systems
				Case HotAir, Hydrogen, Helium
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Useful Static Lift (lbs)", .Lift, wdDouble, "Lift")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					
					
				Case ContraGravGenerator
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Lift", .DesiredLift, wdDouble, "DesiredLift")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Total Lift", VB6.Format(.Lift, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
			End Select
		End With
	End Sub
	
	Private Sub ShowPropsforInstruments(ByVal component As Short, ByVal Key As String)
		Dim m_oCurrentVeh As Object
		'///////////////////////////////////////////
		'Instruments and Electronics
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With m_oCurrentVeh.Components(Key)
			
			Select Case component
				' Fill the window with properties for the correct Collection item
				Case RadioDirectionFinder, RadioCommunicator, TightBeamRadio, VLFRadio, CellularPhone, CellularPhonewithRadio, RadioJammer, ElfReceiver, LaserCommunicator, NeutrinoCommunicator, GravityRippleCommunicator
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Desired Range", .DesiredRange, wdList, "DesiredRange", "short", "medium", "long", "very long", "extreme")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Sensitivity", .Sensitivity, wdList, "Sensitivity", "normal", "sensitive", "very sensitive")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("FTL", .FTL, wdBool, "FTL")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Receive Only", .ReceiveOnly, wdBool, "ReceiveOnly")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Scrambler", .Scrambler, wdBool, "Scrambler")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .FTL = False Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Actual Range", VB6.Format(.Range, "standard") & " miles", wdText, "Disabled")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Actual Range", VB6.Format(.Range, "standard") & " parsecs", wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case Headlight, Searchlight, InfraredSearchlight
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside")
					If component = Headlight Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Range (yards)", .Range, wdDouble, "Range")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Range (miles)", .Range, wdDouble, "Range")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case AstronomicalInstruments, Telescope, lightAmplification, LowlightTV, ExtendableSensorPeriscope
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside")
					If component = lightAmplification Then
					ElseIf component = ExtendableSensorPeriscope Then 
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Periscope Length", .Length, wdDouble, "Length")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Magnification", .Magnification, wdDouble, "Magnification")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case Radar, Ladar, NavigationalRadar, AntiCollisionRadar, AESA, LowResImagingRadar, HiResImagingRadar
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Range", .Range, wdDouble, "Range")
					If component = NavigationalRadar Then
					ElseIf component = AntiCollisionRadar Then 
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("No Targeting", .NoTargeting, wdBool, "NoTargeting")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Search Optimization", .SearchOption, wdList, "SearchOption", "none", "surface search", "air search")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("FTL Option", .FTL, wdBool, "FTL")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Scan Rating", .ScanRating, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case ActiveSonar, PassiveSonar
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					If component = ActiveSonar Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Range", .Range, wdDouble, "Range")
					If component = ActiveSonar Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Active / Passive?", .ActivePassive, wdBool, "ActivePassive")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Depth Finding?", .DepthFinding, wdBool, "DepthFinding")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Dipping Sonar?", .DippingSonar, wdBool, "DippingSonar")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("No Targeting?", .NoTargeting, wdBool, "NoTargeting")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Dipping Sonar?", .DippingSonar, wdBool, "DippingSonar")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Towed Array?", .TowedArray, wdBool, "TowedArray")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Scan Rating", .ScanRating, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case PassiveInfrared, Thermograph, PassiveRadar, PESA
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Range", .Range, wdDouble, "Range")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Scan Rating", .ScanRating, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case Geophone, MAD, MultiScanner, ChemScanner, RadScanner, BioScanner, GravScanner
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Range", .Range, wdDouble, "Range")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Scan Rating", .ScanRating, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case RangingSoundDetector, SurveillanceSoundDetector
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Sensitivity Level", .Level, wdNumber, "Level")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case MeteorologicalInstruments, LowResPlanetarySurveyArray, MedResPlanetarySurveyArray, HighResPlanetarySurveyArray
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Direction", .Direction, wdList, "Direction", "front", "back", "right", "left", "top", "underside")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case SoundSystem, FlightRecorder
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case VehicleCamera, DigitalVehicleCamera, ReconCamera, DigitalReconCamera
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Low light?", .Lowlight, wdBool, "Lowlight")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Infrared?", .Infrared, wdBool, "Infrared")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case NavigationInstruments, AutoPilot, IFF, Transponder, INS, GPS, MilitaryGPS, TFR
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					If component = NavigationInstruments Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Precision?", .Precision, wdBool, "Precision")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case ImprovedOpticalBombSight, AdvancedOpticalBombSight, OpticalBombSight, FireDirectionCenter, HUDWAC, PupilHUDWAC, LaserRangeFinder, LaserDesignator, LaserSpotTracker
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					If component = LaserDesignator Or component = LaserRangeFinder Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Range", .Range, wdDouble, "Range")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case RadarDetector, LaserSensor, LaserRadarDetector, AreaRadarJammer, DeceptiveRadarJammer, InfraredJammer, ChaffDecoyDischarger, SmokeDecoyDischarger, FlareDecoyDischarger, SonarDecoyDischarger, HotSmokeDecoyDischarger, PrismDecoyDischarger, BlackOutGasDecoyDischarger, RadarReflector, BlipEnhancer, TEMPEST
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					Select Case component
						Case AreaRadarJammer, DeceptiveRadarJammer, InfraredJammer
							'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							AddPCLproperty("Jammer Rating", .JammerRating, wdList, "JammerRating", 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20)
						Case RadarDetector, LaserSensor, LaserRadarDetector
							'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							AddPCLproperty("Advanced Version?", .ADVANCED, wdBool, "advanced")
					End Select
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
				Case DecoyChaff, DecoySmoke, DecoyFlares, DecoySonarDecoy, DecoyHotSmoke, DecoyPrism, DecoyBlackOutGas
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
				Case MacroFrame, MainFrame, MicroFrame, MiniComputer, SmallComputer
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Intelligence", .Intelligence, wdList, "Intelligence", "normal", "dumb", "genius")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Configuration", .Configuration, wdList, "Configuration", "normal", "neural-net", "sentient")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Compact?", .Compact, wdBool, "Compact")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hardened?", .Hardened, wdBool, "Hardened")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("High Capacity?", .HighCapacity, wdBool, "HighCapacity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Dedicated?", .Dedicated, wdBool, "Dedicated")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Robot Brain?", .RobotBrain, wdBool, "RobotBrain")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Complexity", .complexity, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .IQ > 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("IQ", .IQ, wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If .DX > 0 Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("DX", .DX, wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
				Case Terminal
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case DatabaseSoftware
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Size (gigs)", .gigabytes, wdDouble, "Gigabytes")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Complexity", .complexity, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					
				Case CartographySoftware, ComputerNavigationSoftware, DatalinkSoftware, TransmissionProfilingSoftware, HoloventureProgram, PersonalitySimulationSoftwareFull, PersonalitySimulationLimited, RoutineVehicleOperationSoftwarePilot, RoutineVehicleOperationSoftwareOther
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Complexity", .complexity, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					
					
				Case FireDirectionSoftware, TargetingSoftware, DamageControlSoftware, GunnerSoftware, RobotSkillProgramsPhysical, RobotSkillProgramsMental
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Bonus Skill", .BonusSkillPoints, wdNumber, "BonusSkillPoints")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Total Skill Points", .SkillPoints, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Complexity", .complexity, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					
					
				Case SurgicalInterface, InterfaceWeb, AutoInterfaceWeb, SocketInterface, NeuralInductionField
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# of Users", .Quantity, wdNumber, "Quantity")
					If component = SocketInterface Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("DR", .dr, wdNumber, "DR")
						AddPCLproperty("Statistics", "", wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("DR", .dr, wdNumber, "DR")
						AddPCLproperty("Statistics", "", wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					End If
					
				Case DeflectorField
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("PD Bonus", "+" & VB6.Format(.PDBonus), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
					'AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
					'AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
					
				Case ForceScreen, VariableForceScreen
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Screen DR", .ForceDR, wdNumber, "ForceDR")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'AddPCLproperty "Volume", Format(.Volume, "standard") & " cf", wdText, "Disabled"
					'AddPCLproperty "Surface Area", Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled"
					'AddPCLproperty "Hit Points", .HitPoints, wdText, "Disabled"
					
			End Select
		End With
	End Sub
	
	Private Sub ShowPropsForMiscellanous(ByVal component As Short, ByVal Key As String)
		Dim m_oCurrentVeh As Object
		'///////////////////////////////////////////
		'Miscellanous equipment
		'///////////////////////////////////////////
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With m_oCurrentVeh.Components(Key)
			
			Select Case component
				' Fill the window with properties for the correct Collection item
				Case ArmMotor
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("ST", .ST, wdNumber, "ST")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Bad Grip?", .BadGrip, wdBool, "BadGrip")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cheap?", .Cheap, wdBool, "Cheap")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Extendable?", .Extendable, wdBool, "Extendable")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Poor Coordination?", .PoorCoordination, wdBool, "PoorCoordination")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Striker?", .Striker, wdBool, "Striker")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case FireExtinguisherSystem, FullFireSuppressionSystem, CompactFireSuppressionSystem
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
				Case BilgePump
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case CompleteWorkshop, MechanicWorkshop, EngineeringWorkshop, ElectronicsWorkshop, ArmouryWorkshop, CompleteMiniWorkshop, ScienceLab, MiniMechanicWorkshop, MiniElectronicsWorkshop, MiniEngineeringWorkshop, MiniArmouryWorkshop
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					If component = ScienceLab Then
						'AddPCLproperty "Skill", .Skill, wdList, "Skill", "astronomy", "biochemistry", "biology", "botany", "chemistry", "computer programming", "criminology", "ecology", "economics", "electronics", "engineering", "forensics", "genetics", "geology", "history", "linguistics", "literature", "mathematics", "metallurgy", "meteorology", "nuclear physics", "occultism", "physics", "physiology", "prospecting", "psychology", "research", "theology", "zoology"
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Skill", .Skill, wdText, "Skill")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case ExtendableLadder, Crane, Winch, PowerShovel, WreckingCrane, ForkLift, VehicularBridge, LaunchCatapult, SkyHook, Bore, SuperBore, EnergyDrill, TractorBeam, PressorBeam, CombinationBeam, CraneWithElectroMagnet
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					Select Case component
						Case ExtendableLadder, Crane, CraneWithElectroMagnet, WreckingCrane
							'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							AddPCLproperty("Crane Height (ft)", .Height, wdNumber, "Height")
						Case PowerShovel, Winch, ForkLift, TractorBeam, PressorBeam, CombinationBeam
							'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							AddPCLproperty("ST", .ST, wdNumber, "ST")
						Case VehicularBridge
							'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							AddPCLproperty("Length (yds)", .Length, wdDouble, "Length")
							'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							AddPCLproperty("Max Supported Weight", .DesiredWeight, wdDouble, "DesiredWeight")
						Case Bore, SuperBore
							'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							AddPCLproperty("Tunneling Ability Per Hour (cf)", .TunnelingAbility, wdDouble, "TunnelingAbility")
						Case SkyHook, LaunchCatapult
					End Select
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case OperatingRoom, StretcherPallet, EmergencySupportUnit, EmergencylightsandSiren, CryonicCapsule, Automed, DiagnosisTable
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					If component = OperatingRoom Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("# of Operating Tables", .OperatingTables, wdNumber, "OperatingTables")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case Stage, Hall, BarRoom, ConferenceRoom, MovieScreenandProjector, MovieScreenandProjectorSmall, HoloventureZone
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					Select Case component
						Case Stage, Hall, BarRoom, ConferenceRoom, HoloventureZone
							'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							AddPCLproperty("Floor Area", .FloorArea, wdDouble, "FloorArea")
						Case Else
							'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					End Select
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					'Note: door and hatch have been remove below.  User should enter these in as "Details" in the options dialog
				Case CargoRamp, Airlock, MembraneAirlock, Forcelock, PassageTube, ArmoredPassageTube
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					Select Case component
						Case Airlock, MembraneAirlock
							'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							AddPCLproperty("# People Supported", .Rating, wdNumber, "Rating")
						Case Else
					End Select
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case TeleportProjector
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# Hexes", .HexCapacity, wdNumber, "HexCapacity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case BrigsandRestraints, BurglarAlarm, HighSecurityAlarm, MutableLicensePlate, OilSprayer, PaintSprayer, SmokeScreen, SpikeDropper
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case VehicleBay, HangerBay, DryDock, SpaceDock, ExternalCradle
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					If component = ExternalCradle Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Total Craft Weight", .CraftWeight, wdDouble, "CraftWeight")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Total Craft Weight", .CraftWeight, wdDouble, "CraftWeight")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Cubic Feet of Craft", .CubicFeetCraft, wdDouble, "CubicFeetCraft")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case ArrestorHook, VehicularParachute
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					If component = VehicularParachute Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Rated Weight", .RatedWeight, wdDouble, "RatedWeight")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case RefuellingProbe, RefuellingDrogue, FuelElectrolysisSystem, HydrogenFuelScoop, AtmosphereProcessor
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					If (component = FuelElectrolysisSystem) Or (component = AtmosphereProcessor) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Processing Capacity (gallons)", .capacity, wdDouble, "Capacity")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case NuclearDamper
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Field Radius (mi)", .Radius, wdDouble, "Radius")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case SmallRealityStabilizer, MediumRealityStabilizer, HeavyRealityStabilizer
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case ModularSocket
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Rated Volume", .RatedVolume, wdDouble, "RatedVolume")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					
					
				Case Module_Renamed
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Waste Weight", .WasteWeight, wdDouble, "WasteWeight")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Waste Volume", .WasteVolume, wdDouble, "WasteVolume")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					
					
					
			End Select
		End With
	End Sub
	
	Private Sub ShowPropsForPowerandFuel(ByVal component As Short, ByVal Key As String)
		Dim m_oCurrentVeh As Object
		'///////////////////////////////////////////
		'Power and Fuel
		'//////////////////////////////////////////
		
		
		Dim listarray() As String
		ReDim listarray(1)
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With m_oCurrentVeh.Components(Key)
			
			' Fill the window with properties for the correct Collection item
			
			Select Case component
				
				Case MuscleEngine
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Maximum Output", .MaxOutput, wdDouble, "MaxOutPut")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Combined Operator ST", .CombinedST, wdNumber, "CombinedST")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Output", VB6.Format(.Output, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Consumed", VB6.Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Remaining", VB6.Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case EarlySteamEngine, ForcedDraftSteamEngine, TripleExpansionSteamEngine, SteamTurbine
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Output", .DesiredOutput, wdDouble, "DesiredOutput")
					If component = SteamTurbine Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Fuel Type", .Fueltype, wdList, "FuelType", "coal", "diesel fuel")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Fuel Type", .Fueltype, wdList, "FuelType", "coal", "wood")
					End If
					
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Output", VB6.Format(.Output, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Consumed", VB6.Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Remaining", VB6.Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Fuel Consumption", .FuelConsumption, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case GasolineEngine, HPGasolineEngine, TurboGasolineEngine, SuperGasolineEngine, TurboHPGasolineEngine, SuperHPGasolineEngine, StandardDieselEngine, TurboStandardDieselEngine, MarineDieselEngine, HPDieselEngine, TurboHPDieselEngine, CeramicEngine, TurboCeramicEngine, SuperCeramicEngine, HPCeramicEngine, TurboHPCeramicEngine, SuperHPCeramicEngine, HydrogenCombustionEngine
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Output", .UnModifiedOutput, wdDouble, "UnModifiedOutput")
					'hydrogen combustion
					If component = HydrogenCombustionEngine Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Fuel Type", .Fueltype, wdList, "FuelType", "hydrogen")
						'aviation fuels
					ElseIf (component = HPCeramicEngine) Or (component = TurboHPCeramicEngine) Or (component = SuperHPCeramicEngine) Or (component = HPGasolineEngine) Or (component = TurboHPGasolineEngine) Or (component = SuperHPGasolineEngine) Then 
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Fuel Type", .Fueltype, wdList, "FuelType", "aviation gas")
						'multifuels
					ElseIf (component = CeramicEngine) Or (component = TurboCeramicEngine) Or (component = SuperCeramicEngine) Then 
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Fuel Type", .Fueltype, wdList, "FuelType", "gasoline", "diesel fuel", "aviation gas", "ethanol", "methanol")
						'diesels with alcohol / propane potential
					ElseIf (component = TurboStandardDieselEngine) Or (component = TurboHPDieselEngine) Or (component = MarineDieselEngine) Or (component = StandardDieselEngine) Or (component = HPDieselEngine) Then 
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Fuel Type", .Fueltype, wdList, "FuelType", "diesel fuel", "propane", "ethanol", "methanol")
						'gasolines with alcohol / propane potential
					ElseIf (component = GasolineEngine) Or (component = TurboGasolineEngine) Or (component = SuperGasolineEngine) Then 
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Fuel Type", .Fueltype, wdList, "FuelType", "gasoline", "propane", "ethanol", "methanol")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Output", VB6.Format(.Output, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Consumed", VB6.Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Remaining", VB6.Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Fuel Consumption", .FuelConsumption, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
					
				Case FuelCell, HPGasTurbine, StandardMHDTurbine, HPMHDTurbine
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Output", .DesiredOutput, wdDouble, "DesiredOutPut")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Closed Cycle?", .ClosedCycle, wdBool, "ClosedCycle")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Output", VB6.Format(.Output, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Consumed", VB6.Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Remaining", VB6.Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Fuel Consumption", .FuelConsumption, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("LOX Consumption", .LOXConsumption, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
				Case StandardGasTurbine, OptimizedGasTurbine
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Output", .DesiredOutput, wdDouble, "DesiredOutPut")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Fuel Type", .Fueltype, wdList, "FuelType", "gasoline", "diesel fuel", "alcohol", "aviation gas", "jet fuel")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Closed Cycle?", .ClosedCycle, wdBool, "ClosedCycle")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Output", VB6.Format(.Output, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Consumed", VB6.Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Remaining", VB6.Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Fuel Consumption", .FuelConsumption, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("LOX Consumption", .LOXConsumption, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
					'only difference between FissionReactor and the others is the Uranium Fuel Rods
				Case FissionReactor
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Output", .DesiredOutput, wdDouble, "DesiredOutput")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Endurance", .Endurance & " yrs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Output", VB6.Format(.Output, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Consumed", VB6.Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Remaining", VB6.Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Fuel Rods Installed", .FuelConsumption, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Fuel Rod Added Cost", .FuelCost, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
					
				Case RTGReactor, NPU, FusionReactor, AntimatterReactor, TotalConversionPowerPlant, CosmicPowerPlant
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Output", .DesiredOutput, wdDouble, "DesiredOutPut")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Endurance", .Endurance & " yrs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Output", VB6.Format(.Output, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Consumed", VB6.Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Remaining", VB6.Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
					
				Case Soulburner, ElementalFurnace, ManaEngine
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Output", .DesiredOutput, wdDouble, "DesiredOutPut")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost for Magic", .MagicCost, wdDouble, "MagicCost")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Output", VB6.Format(.Output, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Consumed", VB6.Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Remaining", VB6.Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case Carnivore, Herbivore, Omnivore, Vampire
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Output", .DesiredOutput, wdDouble, "DesiredOutPut")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Output", VB6.Format(.Output, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Consumed", VB6.Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Remaining", VB6.Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case ClockWork
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Stored Capacity (kWs)", .DesiredOutput, wdDouble, "DesiredOutPut")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Powered Rewind Mechanism?", .PoweredRewinder, wdBool, "PoweredRewinder")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Output", VB6.Format(.Output, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Consumed", VB6.Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Remaining", VB6.Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Rewind Motor ST", .MotorST, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case LeadAcidBattery, AdvancedBattery, Flywheel, RechargeablePowerCell, PowerCell
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Stored Capacity (kWs)", .DesiredOutput, wdDouble, "DesiredOutPut")
					If (component = PowerCell) Or (component = RechargeablePowerCell) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Cell Type", .CellType, wdList, "CellType", "custom", "AA", "A", "B", "C", "D", "E")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Output", VB6.Format(.Output, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Consumed", VB6.Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Remaining", VB6.Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case AntiMatterBay
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Capacity (grams)", .capacity, wdDouble, "Capacity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Failsafe Points", .FailSafePoints, wdNumber, "FailSafePoints")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Fuel Cost", .FuelCost, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Fuel Weight", .FuelWeight, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case StandardTank, lightTank, UltralightTank, StandardSelfSealingTank, lightSelfSealingTank, UltralightSelfSealingTank
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Capacity (gallons)", .capacity, wdDouble, "Capacity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Fuel Type", .Fuel, wdList, "Fuel", "ethanol", "methanol", "aviation gas", "cadmium", "diesel", "gasoline", "jet fuel", "rocket fuel", "water", "hydrogen", "metal/LOX", "oxygen (LOX)", "propane/LNG")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Fire", .Fire, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tank Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tank Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tank Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Fuel Fire", .FuelFire, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Fuel Cost", .FuelCost, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Fuel Weight", .FuelWeight, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case CoalBunker, WoodBunker
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Capacity (cubic ft.)", .capacity, wdDouble, "Capacity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Fuel Cost", .FuelCost, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Fuel Weight", .FuelWeight, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
					'Case Water, Wood, Coal, Gasoline, Diesel, AviationGas, JetFuel, Propane, LiquifiedNaturalGas, EthanolAlchohol, MethanolAlchohol, LiquidHydrogen, LiquidOxygen, Cadmium, MetalLOX, RocketFuel, AntiMatter
					
				Case ElectricContactPower
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Drawn", .DesiredOutput, wdDouble, "DesiredOutPut")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Output", VB6.Format(.Output, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Consumed", VB6.Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Remaining", VB6.Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case LaserBeamedPowerReceiver, MaserBeamedPowerReceiver
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Max Power", .DesiredOutput, wdDouble, "DesiredOutPut")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Output", VB6.Format(.Output, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Consumed", VB6.Format(.PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Remaining", VB6.Format(.Output - .PowerConsumed, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
				Case NitrousOxideBooster
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Max Boost Length", .MaxBoostLength & " seconds", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
				Case Snorkel
					
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillCombustionEngineList)
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Assigned Power Plants", .PowerPlants, wdList, "PowerPlants", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Ruggedized", .Ruggedized, wdBool, "Ruggedized")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
			End Select
			
		End With
	End Sub
	
	Private Sub ShowPropsForArmor(ByVal component As Short, ByVal ComponentsParent As Short, ByVal Key As String)
		Dim m_oCurrentVeh As Object
		Dim listarray() As String
		
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With m_oCurrentVeh.Components(Key)
			
			' Fill the window with properties for the correct Collection item
			
			'///////////////////////////////////////////
			'Armor
			Select Case component
				
				Case ArmorBasicFacing
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillMaterial)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Material", .material, wdList, "Material", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = .FillQuality(.material)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quality", .Quality, wdList, "Quality", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Coating", .coating, wdList, "Coating", "none", "reflective", "retro-reflective")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Radiation Shielding", .radiation, wdBool, "Radiation")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Thermal Superconductor", .thermal, wdBool, "Thermal")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Reactive Armor Plating", .rap, wdBool, "RAP")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Electrified", .electrified, wdBool, "Electrified")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR (Right)", .dr1, wdNumber, "DR1")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR (Left)", .dr2, wdNumber, "DR2")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR (Front)", .dr3, wdNumber, "DR3")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR (Back)", .dr4, wdNumber, "DR4")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR (Top)", .dr5, wdNumber, "DR5")
					If ComponentsParent = Body Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("DR (Bottom)", .dr6, wdNumber, "DR6")
					End If
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Average DR", .AverageDR, wdText, "Disabled") 'MAKE SURE IM Calcing this properly IN THE CLASS depending on whether im dealling with 5 sides or 6!
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Effective DR (Right)", .EffectiveDR1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Effective DR (Left)", .EffectiveDR2, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Effective DR (Front)", .EffectiveDR3, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Effective DR (Back)", .EffectiveDR4, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Effective DR (Top)", .EffectiveDR5, wdText, "Disabled")
					If ComponentsParent = Body Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Effective DR (Bottom)", .EffectiveDR6, wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("PD (Right)", .PD1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("PD (Left)", .PD2, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("PD (Front)", .PD3, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("PD (Back)", .PD4, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("PD (Top)", .PD5, wdText, "Disabled")
					If ComponentsParent = Body Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("PD (Bottom)", .PD6, wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					
					
					
				Case ArmorComplexFacing
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("TL", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Coating", .coating, wdList, "Coating", "none", "reflective", "retro-reflective")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Radiation Shielding", .radiation, wdBool, "Radiation")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Thermal Superconductor", .thermal, wdBool, "Thermal")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Reactive Armor Plating", .rap, wdBool, "RAP")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Electrified", .electrified, wdBool, "Electrified")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillMaterial) 'only needs to be filled once for all sides
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Material (Right)", .material1, wdList, "Material1", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Material (Left)", .material2, wdList, "Material2", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Material (Front)", .material3, wdList, "Material3", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Material (Back)", .material4, wdList, "Material4", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Material (Top)", .material5, wdList, "Material5", listarray)
					If ComponentsParent = Body Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Material (Bottom)", .material6, wdList, "Material6", listarray)
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = .FillQuality(.material1)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quality (Right)", .Quality1, wdList, "Quality1", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = .FillQuality(.material2)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quality (Left)", .Quality2, wdList, "Quality2", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = .FillQuality(.material3)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quality (Front)", .Quality3, wdList, "Quality3", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = .FillQuality(.material4)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quality (Back)", .Quality4, wdList, "Quality4", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = .FillQuality(.material5)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quality (Top)", .Quality5, wdList, "Quality5", listarray)
					If ComponentsParent = Body Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						listarray = .FillQuality(.material6)
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Quality (Bottom)", .Quality6, wdList, "Quality6", listarray)
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR (Right)", .dr1, wdNumber, "DR1", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR (Left)", .dr2, wdNumber, "DR2", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR (Front)", .dr3, wdNumber, "DR3", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR (Back)", .dr4, wdNumber, "DR4", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR (Top)", .dr5, wdNumber, "DR5", listarray)
					If ComponentsParent = Body Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("DR (Bottom)", .dr6, wdNumber, "DR6", listarray)
					End If
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Average DR", .AverageDR, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Effective DR (Right)", .EffectiveDR1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Effective DR (Left)", .EffectiveDR2, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Effective DR (Front)", .EffectiveDR3, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Effective DR (Back)", .EffectiveDR4, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Effective DR (Top)", .EffectiveDR5, wdText, "Disabled")
					If ComponentsParent = Body Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Effective DR (Bottom)", .EffectiveDR6, wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("PD (Right)", .PD1, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("PD (Left)", .PD2, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("PD (Front)", .PD3, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("PD (Back)", .PD4, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("PD (Top)", .PD5, wdText, "Disabled")
					If ComponentsParent = Body Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("PD (Bottom)", .PD6, wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					
					
				Case ArmorComponent
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillMaterial)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Material", .material, wdList, "Material", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = .FillQuality(.material)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quality", .Quality, wdList, "Quality", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("PD", .PD, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					
					
				Case ArmorLocation
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillMaterial)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Material", .material, wdList, "Material", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = .FillQuality(.material)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quality", .Quality, wdList, "Quality", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Coating", .coating, wdList, "Coating", "none", "reflective", "retro-reflective")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Radiation Shielding", .radiation, wdBool, "Radiation")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Thermal Superconductor", .thermal, wdBool, "Thermal")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Reactive Armor Plating", .rap, wdBool, "RAP")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Electrified", .electrified, wdBool, "Electrified")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					If ComponentsParent = Body Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("PD (Right)", .PD1, wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("PD (Left)", .PD2, wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("PD (Front)", .PD3, wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("PD (Back)", .PD4, wdText, "Disabled")
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("PD (Top)", .PD5, wdText, "Disabled")
					ElseIf ComponentsParent = Turret Or ComponentsParent = Popturret Then 
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("PD", .PD6, wdText, "Disabled")
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("PD", .PD, wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					
				Case ArmorOverall, ArmorWheelGuard, ArmorGunShield
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillMaterial)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Material", .material, wdList, "Material", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = .FillQuality(.material)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quality", .Quality, wdList, "Quality", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Coating", .coating, wdList, "Coating", "none", "reflective", "retro-reflective")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Radiation Shielding", .radiation, wdBool, "Radiation")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Thermal Superconductor", .thermal, wdBool, "Thermal")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Reactive Armor Plating", .rap, wdBool, "RAP")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Electrified", .electrified, wdBool, "Electrified")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("PD", .PD, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					
				Case ArmorOpenFrame
					
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = VB6.CopyArray(.FillMaterial)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Material", .material, wdList, "Material", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					listarray = .FillQuality(.material)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quality", .Quality, wdList, "Quality", listarray)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("PD", .PD, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					
			End Select
		End With
	End Sub
	
	Private Sub ShowPropsForMannedVehicles(ByVal component As Short, ByVal Key As String)
		Dim frmNotes As Object
		Dim m_oCurrentVeh As Object
		'//////////////////////////////////////////////
		'Manned Vehicle Components
		'//////////////////////////////////////////////
		'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		With m_oCurrentVeh.Components(Key)
			
			' Fill the window with properties for the correct Collection item
			
			Select Case component
				Case PrimitiveManeuverControl
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					
					
				Case ElectronicDivingControl, ComputerizedDivingControl, MechanicalManeuverControl, ElectronicManeuverControl, ComputerizedManeuverControl, MechanicalDivingControl
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Duplicate?", .duplicate, wdBool, "Duplicate")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					
					
				Case BattlesuitSystem
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Pilot Weight", .PilotWeight, wdDouble, "PilotWeight")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume1", VB6.Format(.Volume1, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case FormFittingBattleSuitSystem
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Pilot Weight", .PilotWeight, wdDouble, "PilotWeight")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight (w/out Pilot)", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Body Volume", VB6.Format(.Volume1, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Turret Volume", VB6.Format(.Volume2, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Arm Volume (each)", VB6.Format(.Volume3, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Leg Volume (each)", VB6.Format(.Volume4, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case CrampedSeat, NormalSeat, RoomySeat
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Exposed?", .Exposed, wdBool, "Exposed")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("G-Seat?", .GSeat, wdBool, "GSeat")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
				Case CycleSeat
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
				Case CrampedStandingRoom, NormalStandingRoom, RoomyStandingRoom
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Exposed?", .Exposed, wdBool, "Exposed")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
				Case Hammock, Bunk, SmallGalley
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Added Volume", .AddedVolume, wdDouble, "AddedVolume")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
				Case Cabin, LuxuryCabin, Suite, LuxurySuite
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Occupancy", .Occupancy, wdNumber, "Occupancy")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("G-Seats?", .GSeat, wdBool, "Gseat")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Added Volume", .AddedVolume, wdDouble, "AddedVolume")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
				Case CrampedCrewStation, NormalCrewStation, RoomyCrewStation
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object frmNotes. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Assignment", frmNotes, wdObject, "StationFunction")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Bridge Access Space?", .BridgeAccessSpace, wdBool, "BridgeAccessSpace")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Exposed?", .Exposed, wdBool, "Exposed")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("G-Seat?", .GSeat, wdBool, "GSeat")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
				Case CycleCrewStation, HarnessCrewStation
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object frmNotes. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Assignment", frmNotes, wdObject, "StationFunction")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
				Case ArtificialGravityUnit
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case EnvironmentalControl, NBCKit, FullLifeSystem, TotalLifeSystem
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# People", .People, wdNumber, "People")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case LimitedLifeSystem
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# People", .People, wdNumber, "People")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# Man Days", .ManDays, wdDouble, "ManDays")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case EjectionSeat, CrewEscapeCapsule, Airbag, CrashWeb, WombTank, GravityWeb
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					If component = CrewEscapeCapsule Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Max Occupancy", .Occupancy, wdNumber, "Occupancy")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					If component = GravityWeb Then
						'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					End If
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
					
				Case GravCompensator
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Tech level", .TL, wdList, "TL", 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16)
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Quantity", .Quantity, wdNumber, "Quantity")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Power Consumption", VB6.Format(.PowerReqt, "standard") & " kW", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("G reduction", .GReduction, wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
				Case Provisions
					AddPCLproperty("Settings", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("# days worth", .occupancydays, wdNumber, "occupancydays")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Settings", .Setting, wdList, "Setting", "auto", "light", "heavy")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("DR", .dr, wdNumber, "DR")
					AddPCLproperty("Statistics", "", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Cost", "$" & VB6.Format(.Cost, "standard"), wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Weight", VB6.Format(.Weight, "standard") & " lbs", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Volume", VB6.Format(.Volume, "standard") & " cf", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Surface Area", VB6.Format(.SurfaceArea, "standard") & " sq ft", wdText, "Disabled")
					'UPGRADE_WARNING: Couldn't resolve default property of object m_oCurrentVeh.Components. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					AddPCLproperty("Hit Points", .HitPoints, wdText, "Disabled")
					
			End Select
		End With
	End Sub
End Module