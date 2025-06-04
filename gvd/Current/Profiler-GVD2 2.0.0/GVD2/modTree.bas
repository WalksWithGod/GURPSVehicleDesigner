Attribute VB_Name = "modTree"
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function LockWindowUpdate Lib "user32" _
    (ByVal hwndLock As Long) As Long

' This function adds the treenodes used to represent the vehicle heiarchy
Public Function GraphVehicle(ByRef oTree As TreeX, ByVal hParentTreeNode As Long, ByVal hNewNodePtr As Long) As Long
vbwProfiler.vbwProcIn 477
vbwProfiler.vbwExecuteLine 10117
On Error GoTo err
vbwProfiler.vbwExecuteLine 10118
    Const LNG_LENGTH = 4
    Dim hChild As Long
    Dim sName As String
    Dim sImage As String
    Dim lngImageIndex As Long

    Dim oNode As cINode

    ' first get local instance of our root object and add it to the tree
    'todo: i forget,his hComponent specifically looking for cIComponent handle?  oArmor doesnt implement this though so it cant be right?  If so
    ' rename that to hItem or something less confusing!
vbwProfiler.vbwExecuteLine 10119
    LockWindowUpdate oTree.hwnd ' prevent repainting til all our nodes are added
vbwProfiler.vbwExecuteLine 10120
    CopyMemory oNode, hNewNodePtr, LNG_LENGTH
vbwProfiler.vbwExecuteLine 10121
    If Not oNode Is Nothing Then
vbwProfiler.vbwExecuteLine 10122
        sName = oNode.Name
vbwProfiler.vbwExecuteLine 10123
        sImage = oNode.Image

vbwProfiler.vbwExecuteLine 10124
        lngImageIndex = AddIconToImageList(App.Path & sImage)
        ' note parent key of "" means its root treeview node
vbwProfiler.vbwExecuteLine 10125
        hParentTreeNode = AddNewChildNode(oTree, sName, App.Path & sImage, hNewNodePtr, hParentTreeNode, True)

        ' todo: add argument to this routine whether to auto expand and select last item
        ' for now we just do it anyway
vbwProfiler.vbwExecuteLine 10126
        oTree.ExpandItem(hParentTreeNode) = True
vbwProfiler.vbwExecuteLine 10127
        oTree.SelectItem(hParentTreeNode) = True

        ' call the recursive subroutine to graph the child components
vbwProfiler.vbwExecuteLine 10128
        Call GraphChildren(oTree, hParentTreeNode, hNewNodePtr)
vbwProfiler.vbwExecuteLine 10129
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
vbwProfiler.vbwExecuteLine 10130
        GraphVehicle = hParentTreeNode
vbwProfiler.vbwExecuteLine 10131
        LockWindowUpdate 0
vbwProfiler.vbwProcOut 477
vbwProfiler.vbwExecuteLine 10132
        Exit Function
    Else
vbwProfiler.vbwExecuteLine 10133 'B
vbwProfiler.vbwExecuteLine 10134
        Debug.Print "modTree:GraphVehicle() -- Error: Invalid interface.  Cannot cast to cINode or valid cIComponent handle not passed in."
    End If
vbwProfiler.vbwExecuteLine 10135 'B
err:
vbwProfiler.vbwExecuteLine 10136
    LockWindowUpdate 0
vbwProfiler.vbwExecuteLine 10137
    If Not oNode Is Nothing Then
vbwProfiler.vbwExecuteLine 10138
        CopyMemory oNode, 0&, LNG_LENGTH
    End If
vbwProfiler.vbwExecuteLine 10139 'B
vbwProfiler.vbwExecuteLine 10140
    If err.Number <> 91 Then
vbwProfiler.vbwExecuteLine 10141
         Debug.Print "frmDesigner:GraphVehicle() -- Error # " & err.Number & " " & err.Description
    End If
vbwProfiler.vbwExecuteLine 10142 'B
vbwProfiler.vbwExecuteLine 10143
    GraphVehicle = False
vbwProfiler.vbwProcOut 477
vbwProfiler.vbwExecuteLine 10144
End Function

Private Sub GraphChildren(ByRef oTree As TreeX, ByVal hParentTreeNode As Long, ByVal hNewNodePtr As Long)
vbwProfiler.vbwProcIn 478
    Dim sName As String
    Dim sImage As String
    Dim sImageKey As String
    Dim hChild As Long
    Dim lngChildCount As Long
    Dim i As Long

    Dim oParent As cINode
    Dim hNode As Long
    Dim oChild As cINode
vbwProfiler.vbwExecuteLine 10145
    Const LNG_LENGTH = 4
    Dim lRet As Long

vbwProfiler.vbwExecuteLine 10146
    On Error GoTo err

    ' this is the recursive function (parent handle is passed in)
vbwProfiler.vbwExecuteLine 10147
    CopyMemory oParent, hNewNodePtr, LNG_LENGTH
vbwProfiler.vbwExecuteLine 10148
    lngChildCount = oParent.childCount

    'Set oChild = oParent.getFirstChild
vbwProfiler.vbwExecuteLine 10149
    For i = 0 To lngChildCount - 1
vbwProfiler.vbwExecuteLine 10150
        Set oChild = oParent.getChild(i)
vbwProfiler.vbwExecuteLine 10151
        hChild = oChild.Handle
vbwProfiler.vbwExecuteLine 10152
        sName = oChild.Name
vbwProfiler.vbwExecuteLine 10153
        sImage = oChild.Image
vbwProfiler.vbwExecuteLine 10154
        sImageKey = App.Path & sImage
vbwProfiler.vbwExecuteLine 10155
        lRet = AddIconToImageList(sImageKey)
vbwProfiler.vbwExecuteLine 10156
        hNode = AddNewChildNode(oTree, sName, sImageKey, hChild, hParentTreeNode, True)

        ' todo: add argument to this routine whether to auto expand and select last item
        ' for now we just do it anyway
vbwProfiler.vbwExecuteLine 10157
        oTree.ExpandItem(hNode) = True
vbwProfiler.vbwExecuteLine 10158
        oTree.SelectItem(hNode) = True

        ' call recurse(hChild) to load a node for each child it may have
vbwProfiler.vbwExecuteLine 10159
        Call GraphChildren(oTree, hNode, hChild)
vbwProfiler.vbwExecuteLine 10160
    Next
vbwProfiler.vbwExecuteLine 10161
    Set oChild = Nothing
vbwProfiler.vbwExecuteLine 10162
    CopyMemory oParent, 0&, LNG_LENGTH
vbwProfiler.vbwProcOut 478
vbwProfiler.vbwExecuteLine 10163
    Exit Sub
err:
vbwProfiler.vbwExecuteLine 10164
    If Not oParent Is Nothing Then
vbwProfiler.vbwExecuteLine 10165
        CopyMemory oParent, 0&, LNG_LENGTH
    End If
vbwProfiler.vbwExecuteLine 10166 'B
vbwProfiler.vbwExecuteLine 10167
    Set oChild = Nothing
vbwProfiler.vbwProcOut 478
vbwProfiler.vbwExecuteLine 10168
End Sub

Function AddNewChildNode(ByRef oTree As TreeX, ByVal NodeName As String, ByRef sImage As String, ByVal hHandle As Long, ByVal hParent As Long, ByVal bDown As Boolean) As Long
    'Add a node using tvwChild
vbwProfiler.vbwProcIn 479
    Dim iIndex As Integer
    Dim lRet As Long

vbwProfiler.vbwExecuteLine 10169
    On Error GoTo myerr 'in case the treeview does not have a node selected

vbwProfiler.vbwExecuteLine 10170
    If bDown Then
vbwProfiler.vbwExecuteLine 10171
        lRet = oTree.AddItemImage(NodeName, sImage, hParent, True)
    Else
vbwProfiler.vbwExecuteLine 10172 'B
vbwProfiler.vbwExecuteLine 10173
        lRet = oTree.AddItemImage(NodeName, sImage, hParent, False)
    End If
vbwProfiler.vbwExecuteLine 10174 'B
vbwProfiler.vbwExecuteLine 10175
    oTree.ItemData(lRet) = hHandle
vbwProfiler.vbwExecuteLine 10176
    AddNewChildNode = lRet
vbwProfiler.vbwProcOut 479
vbwProfiler.vbwExecuteLine 10177
    Exit Function

myerr:
vbwProfiler.vbwExecuteLine 10178
    Debug.Print "ModMain:AddNewChildNode() -- Error #" & err.Number & " " & err.Description
    'Display a messge telling the user to select a node
vbwProfiler.vbwExecuteLine 10179
    MsgBox ("You must select a Node to add a child node" & vbCrLf _
       & "If the TreeView is empty, you must first create a new vehicle")
vbwProfiler.vbwProcOut 479
vbwProfiler.vbwExecuteLine 10180
    Exit Function
vbwProfiler.vbwProcOut 479
vbwProfiler.vbwExecuteLine 10181
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
vbwProfiler.vbwProcIn 480
vbwProfiler.vbwProcOut 480
vbwProfiler.vbwExecuteLine 10182
End Sub


