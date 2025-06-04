Option Strict Off
Option Explicit On
Module modTree
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	Private Declare Sub CopyMemory Lib "kernel32"  Alias "RtlMoveMemory"(ByRef hpvDest As Any, ByRef hpvSource As Any, ByVal cbCopy As Integer)
	Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Integer) As Integer
	
	' This function adds the treenodes used to represent the vehicle heiarchy
	Public Function GraphVehicle(ByRef oTree As TreeX, ByVal hParentTreeNode As Integer, ByVal hNewNodePtr As Integer) As Integer
		On Error GoTo err_Renamed
		Const LNG_LENGTH As Short = 4
		Dim hChild As Integer
		Dim sName As String
		Dim sImage As String
		Dim lngImageIndex As Integer
		
		Dim oNode As _cINode
		
		' first get local instance of our root object and add it to the tree
		'todo: i forget,his hComponent specifically looking for cIComponent handle?  oArmor doesnt implement this though so it cant be right?  If so
		' rename that to hItem or something less confusing!
		'UPGRADE_WARNING: Couldn't resolve default property of object oTree.hwnd. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		LockWindowUpdate(oTree.hwnd) ' prevent repainting til all our nodes are added
		'UPGRADE_WARNING: Couldn't resolve default property of object oNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(oNode, hNewNodePtr, LNG_LENGTH)
		If Not oNode Is Nothing Then
			sName = oNode.Name
			sImage = oNode.Image
			
			lngImageIndex = AddIconToImageList(My.Application.Info.DirectoryPath & sImage)
			' note parent key of "" means its root treeview node
			hParentTreeNode = AddNewChildNode(oTree, sName, My.Application.Info.DirectoryPath & sImage, hNewNodePtr, hParentTreeNode, True)
			
			' todo: add argument to this routine whether to auto expand and select last item
			' for now we just do it anyway
			'UPGRADE_WARNING: Couldn't resolve default property of object oTree.ExpandItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oTree.ExpandItem(hParentTreeNode) = True
			'UPGRADE_WARNING: Couldn't resolve default property of object oTree.SelectItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oTree.SelectItem(hParentTreeNode) = True
			
			' call the recursive subroutine to graph the child components
			Call GraphChildren(oTree, hParentTreeNode, hNewNodePtr)
			'UPGRADE_WARNING: Couldn't resolve default property of object oNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(oNode, 0, LNG_LENGTH) ' NOTE: Important to release this thusly, or it will crash the IDE
			
			' now we still need to graph our NON child components such as
			' hull, surface, profiles, crew, and armor
			
			' Since this modTree module app indepedant [its also used by the rule wizard] it
			' only cares about the iNode hiearchy and the treex reference being passed in
			' we want every cInode auto graphed without special treatment
			' Problem is, our rule editor doesnt use armor objects, or surface, or hull, crew, etc...
			' Well, this is inconsequential since we are happy with the component architecture whereby
			' armor, hull, surface, frames,etc are composite objects that do need to be graphed.  If that means
			' addding special cases to this function, so be it.
			
			' this returns the handle to the first node that was added
			GraphVehicle = hParentTreeNode
			LockWindowUpdate(0)
			Exit Function
		Else
			Debug.Print("modTree:GraphVehicle() -- Error: Invalid interface.  Cannot cast to cINode or valid cIComponent handle not passed in.")
		End If
err_Renamed: 
		LockWindowUpdate(0)
		If Not oNode Is Nothing Then
			'UPGRADE_WARNING: Couldn't resolve default property of object oNode. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(oNode, 0, LNG_LENGTH)
		End If
		If Err.Number <> 91 Then Debug.Print("frmDesigner:GraphVehicle() -- Error # " & Err.Number & " " & Err.Description)
		GraphVehicle = False
	End Function
	
	Private Sub GraphChildren(ByRef oTree As TreeX, ByVal hParentTreeNode As Integer, ByVal hNewNodePtr As Integer)
		Dim sName As String
		Dim sImage As String
		Dim sImageKey As String
		Dim hChild As Integer
		Dim lngChildCount As Integer
		Dim i As Integer
		
		Dim oParent As _cINode
		Dim hNode As Integer
		Dim oChild As _cINode
		Const LNG_LENGTH As Short = 4
		Dim lRet As Integer
		
		On Error GoTo err_Renamed
		
		' this is the recursive function (parent handle is passed in)
		'UPGRADE_WARNING: Couldn't resolve default property of object oParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(oParent, hNewNodePtr, LNG_LENGTH)
		lngChildCount = oParent.childCount
		
		'Set oChild = oParent.getFirstChild
		For i = 0 To lngChildCount - 1
			oChild = oParent.getChild(i)
			hChild = oChild.Handle
			sName = oChild.Name
			sImage = oChild.Image
			sImageKey = My.Application.Info.DirectoryPath & sImage
			lRet = AddIconToImageList(sImageKey)
			hNode = AddNewChildNode(oTree, sName, sImageKey, hChild, hParentTreeNode, True)
			
			' todo: add argument to this routine whether to auto expand and select last item
			' for now we just do it anyway
			'UPGRADE_WARNING: Couldn't resolve default property of object oTree.ExpandItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oTree.ExpandItem(hNode) = True
			'UPGRADE_WARNING: Couldn't resolve default property of object oTree.SelectItem. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			oTree.SelectItem(hNode) = True
			
			' call recurse(hChild) to load a node for each child it may have
			Call GraphChildren(oTree, hNode, hChild)
		Next 
		'UPGRADE_NOTE: Object oChild may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oChild = Nothing
		'UPGRADE_WARNING: Couldn't resolve default property of object oParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CopyMemory(oParent, 0, LNG_LENGTH)
		Exit Sub
err_Renamed: 
		If Not oParent Is Nothing Then
			'UPGRADE_WARNING: Couldn't resolve default property of object oParent. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			CopyMemory(oParent, 0, LNG_LENGTH)
		End If
		'UPGRADE_NOTE: Object oChild may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		oChild = Nothing
	End Sub
	
	Function AddNewChildNode(ByRef oTree As TreeX, ByVal NodeName As String, ByRef sImage As String, ByVal hHandle As Integer, ByVal hParent As Integer, ByVal bDown As Boolean) As Integer
		'Add a node using tvwChild
		Dim iIndex As Short
		Dim lRet As Integer
		
		On Error GoTo myerr 'in case the treeview does not have a node selected
		
		If bDown Then
			'UPGRADE_WARNING: Couldn't resolve default property of object oTree.AddItemImage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lRet = oTree.AddItemImage(NodeName, sImage, hParent, True)
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object oTree.AddItemImage. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			lRet = oTree.AddItemImage(NodeName, sImage, hParent, False)
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object oTree.ItemData. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		oTree.ItemData(lRet) = hHandle
		AddNewChildNode = lRet
		Exit Function
		
myerr: 
		Debug.Print("ModMain:AddNewChildNode() -- Error #" & Err.Number & " " & Err.Description)
		'Display a messge telling the user to select a node
		MsgBox("You must select a Node to add a child node" & vbCrLf & "If the TreeView is empty, you must first create a new vehicle")
		Exit Function
	End Function
	
	Sub GetFirstParent() 'todo: obsolete, but i think GVD still tries to call this... resolve it and then
		' delete this sub routine
		'    'Find the first parent node in the TreeView
		'    On Error GoTo myerr
		'    Dim i As Integer
		'    Dim nTmp As Integer
		'    For i = 1 To frmDesigner.treeVehicle.Nodes.Count
		'        'This will give an error if there is no parent
		'        nTmp = frmDesigner.treeVehicle.Nodes(i).Parent.index
		'     Next
		'    Exit Sub
		'myerr:
		'    p_nIndex = i
		'    Exit Sub
	End Sub
End Module