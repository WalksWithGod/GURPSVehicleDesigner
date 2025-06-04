Attribute VB_Name = "modTree"
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function LockWindowUpdate Lib "user32" _
    (ByVal hwndLock As Long) As Long

' This function adds the treenodes used to represent the vehicle heiarchy
Public Function GraphVehicle(ByRef oTree As TreeX, ByVal hParentTreeNode As Long, ByVal hNewNodePtr As Long) As Long
On Error GoTo err
    Const LNG_LENGTH = 4
    Dim hChild As Long
    Dim sName As String
    Dim sImage As String
    Dim lngImageIndex As Long

    Dim oNode As cINode
    
    ' first get local instance of our root object and add it to the tree
    'todo: i forget,his hComponent specifically looking for cIComponent handle?  oArmor doesnt implement this though so it cant be right?  If so
    ' rename that to hItem or something less confusing!
    LockWindowUpdate oTree.hwnd ' prevent repainting til all our nodes are added
    CopyMemory oNode, hNewNodePtr, LNG_LENGTH
    If Not oNode Is Nothing Then
        sName = oNode.Name
        sImage = oNode.Image
             
        lngImageIndex = AddIconToImageList(App.Path & sImage)
        ' note parent key of "" means its root treeview node
        hParentTreeNode = AddNewChildNode(oTree, sName, App.Path & sImage, hNewNodePtr, hParentTreeNode, True)
        
        ' todo: add argument to this routine whether to auto expand and select last item
        ' for now we just do it anyway
        oTree.ExpandItem(hParentTreeNode) = True
        oTree.SelectItem(hParentTreeNode) = True
                
        ' call the recursive subroutine to graph the child components
        Call GraphChildren(oTree, hParentTreeNode, hNewNodePtr)
        CopyMemory oNode, 0&, LNG_LENGTH ' NOTE: Important to release this thusly, or it will crash the IDE
        
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
        LockWindowUpdate 0
        Exit Function
    Else
        Debug.Print "modTree:GraphVehicle() -- Error: Invalid interface.  Cannot cast to cINode or valid cIComponent handle not passed in."
    End If
err:
    LockWindowUpdate 0
    If Not oNode Is Nothing Then
        CopyMemory oNode, 0&, LNG_LENGTH
    End If
    If err.Number <> 91 Then Debug.Print "frmDesigner:GraphVehicle() -- Error # " & err.Number & " " & err.Description
    GraphVehicle = False
End Function

Private Sub GraphChildren(ByRef oTree As TreeX, ByVal hParentTreeNode As Long, ByVal hNewNodePtr As Long)
    Dim sName As String
    Dim sImage As String
    Dim sImageKey As String
    Dim hChild As Long
    Dim lngChildCount As Long
    Dim i As Long
    
    Dim oParent As cINode
    Dim hNode As Long
    Dim oChild As cINode
    Const LNG_LENGTH = 4
    Dim lRet As Long
    
    On Error GoTo err
    
    ' this is the recursive function (parent handle is passed in)
    CopyMemory oParent, hNewNodePtr, LNG_LENGTH
    lngChildCount = oParent.childCount
    
    'Set oChild = oParent.getFirstChild
    For i = 0 To lngChildCount - 1
        Set oChild = oParent.getChild(i)
        hChild = oChild.Handle
        sName = oChild.Name
        sImage = oChild.Image
        sImageKey = App.Path & sImage
        lRet = AddIconToImageList(sImageKey)
        hNode = AddNewChildNode(oTree, sName, sImageKey, hChild, hParentTreeNode, True)
        
        ' todo: add argument to this routine whether to auto expand and select last item
        ' for now we just do it anyway
        oTree.ExpandItem(hNode) = True
        oTree.SelectItem(hNode) = True
    
        ' call recurse(hChild) to load a node for each child it may have
        Call GraphChildren(oTree, hNode, hChild)
    Next
    Set oChild = Nothing
    CopyMemory oParent, 0&, LNG_LENGTH
    Exit Sub
err:
    If Not oParent Is Nothing Then
        CopyMemory oParent, 0&, LNG_LENGTH
    End If
    Set oChild = Nothing
End Sub

Function AddNewChildNode(ByRef oTree As TreeX, ByVal NodeName As String, ByRef sImage As String, ByVal hHandle As Long, ByVal hParent As Long, ByVal bDown As Boolean) As Long
    'Add a node using tvwChild
    Dim iIndex As Integer
    Dim lRet As Long
           
    On Error GoTo myerr 'in case the treeview does not have a node selected
    
    If bDown Then
        lRet = oTree.AddItemImage(NodeName, sImage, hParent, True)
    Else
        lRet = oTree.AddItemImage(NodeName, sImage, hParent, False)
    End If
    oTree.ItemData(lRet) = hHandle
    AddNewChildNode = lRet
    Exit Function
    
myerr:
    Debug.Print "ModMain:AddNewChildNode() -- Error #" & err.Number & " " & err.Description
    'Display a messge telling the user to select a node
    MsgBox ("You must select a Node to add a child node" & vbCrLf _
       & "If the TreeView is empty, you must first create a new vehicle")
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

