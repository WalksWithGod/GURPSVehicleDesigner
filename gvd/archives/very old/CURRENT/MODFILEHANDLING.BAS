Attribute VB_Name = "modFileHandling"
Option Explicit
'======================================================
' Variable declaration for the UDT's
'======================================================
Dim udtBody As Body
Dim udtSkid As Skid
Dim udtWheel As Wheel
Dim udtTrack As Track
Dim udtLeg As Leg
Dim udtWing As Wing
Dim udtRotar As Rotar
Dim udtHydrofoil As Hydrofoil
Dim udtHovercraft As Hovercraft
Dim udtSuperstructure As Superstructure
Dim udtOpenMount As OpenMount
Dim udtMast As Mast
Dim udtPod As Pod
Dim udtTurret As Turret
Dim udtPopturret As Popturret
Dim udtGasbag As Gasbag
Dim udtArm As Arm

Dim udtArtilleryandGun As ArtilleryandGun
Dim udtSettings As Settings

'======================================================
' These are the UDT's for all the Subassemblies
'======================================================
Private Type Wheel
    Parent As String
    Key As String
    Text As String
    Image As Integer
    SelectedImage As Integer
    Name As String
    Number As Integer
    ImprovedSuspension As Boolean
    Wheelblades As Integer
    SnowTires As Boolean
    RacingTires As Boolean
    PunctureResistant As Boolean
    RetractLocation As Integer
    ImprovedBreaks As Boolean
    AllwheelSteering As Boolean
    Smartwheels As Boolean
End Type

Private Type Body
    Name As String
    Text As String
    Key As String
    Parent As String
    Image As Integer
    SelectedImage As Integer
End Type

Private Type Wing
    Name As String
    Text As String
    Key As String
    Parent As String
    Image As Integer
    SelectedImage As Integer
End Type

Private Type Rotar
    Name As String
    Text As String
    Key As String
    Parent As String
    Image As Integer
    SelectedImage As Integer
End Type

Private Type Hydrofoil
    Name As String
    Text As String
    Key As String
    Parent As String
    Image As Integer
    SelectedImage As Integer
End Type

Private Type Hovercraft
    Name As String
    Text As String
    Key As String
    Parent As String
    Image As Integer
    SelectedImage As Integer
End Type

Private Type Skid
    Name As String
    Text As String
    Key As String
    Parent As String
    Image As Integer
    SelectedImage As Integer
End Type

Private Type Track
    Name As String
    Text As String
    Key As String
    Parent As String
    Image As Integer
    SelectedImage As Integer
End Type

Private Type Leg
    Name As String
    Text As String
    Key As String
    Parent As String
    Image As Integer
    SelectedImage As Integer
End Type

Private Type Superstructure
    Name As String
    Text As String
    Key As String
    Parent As String
    Image As Integer
    SelectedImage As Integer
End Type

Private Type OpenMount
    Name As String
    Text As String
    Key As String
    Parent As String
    Image As Integer
    SelectedImage As Integer
End Type

Private Type Mast
    Name As String
    Text As String
    Key As String
    Parent As String
    Image As Integer
    SelectedImage As Integer
End Type

Private Type Pod
    Name As String
    Text As String
    Key As String
    Parent As String
    Image As Integer
    SelectedImage As Integer
End Type

Private Type Turret
    Name As String
    Text As String
    Key As String
    Parent As String
    Image As Integer
    SelectedImage As Integer
End Type

Private Type Popturret
    Name As String
    Text As String
    Key As String
    Parent As String
    Image As Integer
    SelectedImage As Integer
End Type

Private Type Arm
    Name As String
    Text As String
    Key As String
    Parent As String
    Image As Integer
    SelectedImage As Integer
End Type

Private Type Gasbag
    Name As String
    Text As String
    Key As String
    Parent As String
    Image As Integer
    SelectedImage As Integer
End Type



'========================================================
' These are the UDT's for components
'========================================================

Private Type ArtilleryandGun
    Name As String
    Parent As String
    Key As String
    Text As String
    Image As Integer
    SelectedImage As Integer
    Malf As String
    Type As String
    Damage As String
    SS As String
    Acc As String
    halfD As String
    Max As String
    RoF As String
    Wt As String
    Cost As String
    WPS As String
    VPS As String
    CPS As String
    Ldrs As String
End Type

'======================================================
' This is the UDT for the INI file
'======================================================

Type Settings
     InitialDir As String
End Type


Sub LoadRecords(TheFile As String)


' Reads the file and creates collection and builds the tree nodes
Dim i As Integer
Dim ID As Integer

' Destroy any object variables that might already exist in the collection
Set colVehicle = Nothing

' Clear the treeview of any nodes that might already exist
frmDesigner.TreeView1.Nodes.Clear

' Open the file specified using Binary access
Open TheFile For Binary As #1 ' need to replace this hardcoded path

'  Until the end of the file is reached Do the following
i = 1
While i < LOF(1)
    ' read the current positions ID byte to deterine the object's class
    Get #1, i, ID
 
' once the type of class is known, read in all of those values according to that ID

Select Case ID
    Case ArtilleryandGun ' ArtilleryandGun value is 17
        Get #1, , udtArtilleryandGun
        i = Seek(1)
        'Create the object and add it to the collection and the tree
        Set objArtilleryandGun = New clsArtilleryandGun
        colVehicle.Add objArtilleryandGun, udtArtilleryandGun.Key  ' add the new object to the collection
         
        frmDesigner.TreeView1.Nodes.Add udtArtilleryandGun.Parent, tvwChild, udtArtilleryandGun.Key, udtArtilleryandGun.Text, udtArtilleryandGun.Image, udtArtilleryandGun.SelectedImage
        frmDesigner.TreeView1.Nodes.Item(udtArtilleryandGun.Key).EnsureVisible ' expand the tree branch

        With colVehicle.Item(udtArtilleryandGun.Key) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
            .Key = udtArtilleryandGun.Key
            .Parent = udtArtilleryandGun.Parent
            .Text = udtArtilleryandGun.Text
            .Image = udtArtilleryandGun.Image
            .SelectedImage = udtArtilleryandGun.SelectedImage
        End With
 Case Body  ' Body value is 0
        Get #1, , udtBody
        i = Seek(1)
        ' Create the object and add it to the collection and the tree
        Set objBody = New clsBody
        colVehicle.Add objBody, udtBody.Key  ' add the new object to the collection
            If udtBody.Parent = "" Then 'All root nodes have 0_ in the parent field
                    frmDesigner.TreeView1.Nodes.Add , , udtBody.Key, udtBody.Text, udtBody.Image, udtBody.SelectedImage
                Else 'All child nodes will have the parent key stored in the parent field
                    frmDesigner.TreeView1.Nodes.Add udtBody.Parent, tvwChild, udtBody.Key, udtBody.Text, udtBody.Image, udtBody.SelectedImage
            End If
        frmDesigner.TreeView1.Nodes.Item(udtBody.Key).EnsureVisible ' expand the tree branch
        ' select the first node
        frmDesigner.TreeView1.Nodes(1).Selected = True 'Select the first Node
            With colVehicle.Item(udtBody.Key) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
                .Key = udtBody.Key
                '.Parent = frmDesigner.TreeView1.Nodes.Item(udtBody.Key).Parent.Key
                .Text = udtBody.Text
                .Image = udtBody.Image
                .SelectedImage = udtBody.SelectedImage
            End With
 Case Skid ' Skid value is 1
        Get #1, , udtSkid
        i = Seek(1)
        'Create the object and add it to the collection and the tree
        Set objSkid = New clsSkid
        colVehicle.Add objSkid, udtSkid.Key  ' add the new object to the collection
        frmDesigner.TreeView1.Nodes.Add udtSkid.Parent, tvwChild, udtSkid.Key, udtSkid.Text, udtSkid.Image, udtSkid.SelectedImage
        frmDesigner.TreeView1.Nodes.Item(udtSkid.Key).EnsureVisible ' expand the tree branch
            With colVehicle.Item(udtSkid.Key) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
                .Key = frmDesigner.TreeView1.Nodes.Item(udtSkid.Key).Key
                .Parent = frmDesigner.TreeView1.Nodes.Item(udtSkid.Key).Parent.Key
                .Text = frmDesigner.TreeView1.Nodes.Item(udtSkid.Key).Text
                .Image = frmDesigner.TreeView1.Nodes.Item(udtSkid.Key).Image
                .SelectedImage = frmDesigner.TreeView1.Nodes.Item(udtSkid.Key).SelectedImage
            End With
 Case Wheel ' Wheel value is 2
        Get #1, , udtWheel
        i = Seek(1)
        'Create the object and add it to the collection and the tree
        Set objWheel = New clsWheel
        colVehicle.Add objWheel, udtWheel.Key  ' add the new object to the collection
        frmDesigner.TreeView1.Nodes.Add udtWheel.Parent, tvwChild, udtWheel.Key, udtWheel.Text, udtWheel.Image, udtWheel.SelectedImage
        frmDesigner.TreeView1.Nodes.Item(udtWheel.Key).EnsureVisible ' expand the tree branch
        'Update_Object udtWheel.Key
            With colVehicle.Item(udtWheel.Key) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
                .Key = frmDesigner.TreeView1.Nodes.Item(udtWheel.Key).Key
                .Parent = frmDesigner.TreeView1.Nodes.Item(udtWheel.Key).Parent.Key
                .Text = frmDesigner.TreeView1.Nodes.Item(udtWheel.Key).Text
                .Image = frmDesigner.TreeView1.Nodes.Item(udtWheel.Key).Image
                .SelectedImage = frmDesigner.TreeView1.Nodes.Item(udtWheel.Key).SelectedImage
            End With
 Case Track ' Track value is 3
        Get #1, , udtTrack
        i = Seek(1)
        'Create the object and add it to the collection and the tree
        Set objTrack = New clsTrack
        colVehicle.Add objTrack, udtTrack.Key  ' add the new object to the collection
        frmDesigner.TreeView1.Nodes.Add udtTrack.Parent, tvwChild, udtTrack.Key, udtTrack.Text, udtTrack.Image, udtTrack.SelectedImage
        frmDesigner.TreeView1.Nodes.Item(udtTrack.Key).EnsureVisible ' expand the tree branch
        'Update_Object udtTrack.Key
            With colVehicle.Item(udtTrack.Key) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
                .Key = frmDesigner.TreeView1.Nodes.Item(udtTrack.Key).Key
                .Parent = frmDesigner.TreeView1.Nodes.Item(udtTrack.Key).Parent.Key
                .Text = frmDesigner.TreeView1.Nodes.Item(udtTrack.Key).Text
                .Image = frmDesigner.TreeView1.Nodes.Item(udtTrack.Key).Image
                .SelectedImage = frmDesigner.TreeView1.Nodes.Item(udtTrack.Key).SelectedImage
            End With
 Case Leg ' Leg value is 4
        Get #1, , udtLeg
        i = Seek(1)
        'Create the object and add it to the collection and the tree
        Set objLeg = New clsLeg
        colVehicle.Add objLeg, udtLeg.Key  ' add the new object to the collection
        frmDesigner.TreeView1.Nodes.Add udtLeg.Parent, tvwChild, udtLeg.Key, udtLeg.Text, udtLeg.Image, udtLeg.SelectedImage
        frmDesigner.TreeView1.Nodes.Item(udtLeg.Key).EnsureVisible ' expand the tree branch
        'Update_Object udtLeg.Key
            With colVehicle.Item(udtLeg.Key) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
                .Key = frmDesigner.TreeView1.Nodes.Item(udtLeg.Key).Key
                .Parent = frmDesigner.TreeView1.Nodes.Item(udtLeg.Key).Parent.Key
                .Text = frmDesigner.TreeView1.Nodes.Item(udtLeg.Key).Text
                .Image = frmDesigner.TreeView1.Nodes.Item(udtLeg.Key).Image
                .SelectedImage = frmDesigner.TreeView1.Nodes.Item(udtLeg.Key).SelectedImage
            End With
 Case Wing ' Wing value is 5
        Get #1, , udtWing
        i = Seek(1)
        'Create the object and add it to the collection and the tree
        Set objWing = New clsWing
        colVehicle.Add objWing, udtWing.Key  ' add the new object to the collection
        frmDesigner.TreeView1.Nodes.Add udtWing.Parent, tvwChild, udtWing.Key, udtWing.Text, udtWing.Image, udtWing.SelectedImage
        frmDesigner.TreeView1.Nodes.Item(udtWing.Key).EnsureVisible ' expand the tree branch
        'Update_Object udtWing.Key

        With colVehicle.Item(udtWing.Key) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
            .Key = frmDesigner.TreeView1.Nodes.Item(udtWing.Key).Key
            .Parent = frmDesigner.TreeView1.Nodes.Item(udtWing.Key).Parent.Key
            .Text = frmDesigner.TreeView1.Nodes.Item(udtWing.Key).Text
            .Image = frmDesigner.TreeView1.Nodes.Item(udtWing.Key).Image
            .SelectedImage = frmDesigner.TreeView1.Nodes.Item(udtWing.Key).SelectedImage
        End With

Case Rotar ' Rotar value is 6
        
        Get #1, , udtRotar
        i = Seek(1)
        'Create the object and add it to the collection and the tree
        Set objRotar = New clsRotar
        colVehicle.Add objRotar, udtRotar.Key  ' add the new object to the collection
        frmDesigner.TreeView1.Nodes.Add udtRotar.Parent, tvwChild, udtRotar.Key, udtRotar.Text, udtRotar.Image, udtRotar.SelectedImage
        frmDesigner.TreeView1.Nodes.Item(udtRotar.Key).EnsureVisible ' expand the tree branch
        'Update_Object udtRotar.Key

        With colVehicle.Item(udtRotar.Key) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
            .Key = frmDesigner.TreeView1.Nodes.Item(udtRotar.Key).Key
            .Parent = frmDesigner.TreeView1.Nodes.Item(udtRotar.Key).Parent.Key
            .Text = frmDesigner.TreeView1.Nodes.Item(udtRotar.Key).Text
            .Image = frmDesigner.TreeView1.Nodes.Item(udtRotar.Key).Image
            .SelectedImage = frmDesigner.TreeView1.Nodes.Item(udtRotar.Key).SelectedImage
        End With

Case Hydrofoil ' Hydrofoil value is 7
        
        Get #1, , udtHydrofoil
        i = Seek(1)
        'Create the object and add it to the collection and the tree
        Set objHydrofoil = New clsHydrofoil
        colVehicle.Add objHydrofoil, udtHydrofoil.Key  ' add the new object to the collection
        frmDesigner.TreeView1.Nodes.Add udtHydrofoil.Parent, tvwChild, udtHydrofoil.Key, udtHydrofoil.Text, udtHydrofoil.Image, udtHydrofoil.SelectedImage
        frmDesigner.TreeView1.Nodes.Item(udtHydrofoil.Key).EnsureVisible ' expand the tree branch
        'Update_Object udtHydrofoil.Key

        With colVehicle.Item(udtHydrofoil.Key) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
            .Key = frmDesigner.TreeView1.Nodes.Item(udtHydrofoil.Key).Key
            .Parent = frmDesigner.TreeView1.Nodes.Item(udtHydrofoil.Key).Parent.Key
            .Text = frmDesigner.TreeView1.Nodes.Item(udtHydrofoil.Key).Text
            .Image = frmDesigner.TreeView1.Nodes.Item(udtHydrofoil.Key).Image
            .SelectedImage = frmDesigner.TreeView1.Nodes.Item(udtHydrofoil.Key).SelectedImage
        End With

Case Hovercraft ' Hovercraft value is 8
        
        Get #1, , udtHovercraft
        i = Seek(1)
        'Create the object and add it to the collection and the tree
        Set objHovercraft = New clsHovercraft
        colVehicle.Add objHovercraft, udtHovercraft.Key  ' add the new object to the collection
        frmDesigner.TreeView1.Nodes.Add udtHovercraft.Parent, tvwChild, udtHovercraft.Key, udtHovercraft.Text, udtHovercraft.Image, udtHovercraft.SelectedImage
        frmDesigner.TreeView1.Nodes.Item(udtHovercraft.Key).EnsureVisible ' expand the tree branch
        'Update_Object udtHovercraft.Key

        With colVehicle.Item(udtHovercraft.Key) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
            .Key = frmDesigner.TreeView1.Nodes.Item(udtHovercraft.Key).Key
            .Parent = frmDesigner.TreeView1.Nodes.Item(udtHovercraft.Key).Parent.Key
            .Text = frmDesigner.TreeView1.Nodes.Item(udtHovercraft.Key).Text
            .Image = frmDesigner.TreeView1.Nodes.Item(udtHovercraft.Key).Image
            .SelectedImage = frmDesigner.TreeView1.Nodes.Item(udtHovercraft.Key).SelectedImage
        End With

Case Superstructure ' Superstructure value is 9
        
        Get #1, , udtSuperstructure
        i = Seek(1)
        'Create the object and add it to the collection and the tree
        Set objSuperStructure = New clsSuperStructure
        colVehicle.Add objSuperStructure, udtSuperstructure.Key  ' add the new object to the collection
        frmDesigner.TreeView1.Nodes.Add udtSuperstructure.Parent, tvwChild, udtSuperstructure.Key, udtSuperstructure.Text, udtSuperstructure.Image, udtSuperstructure.SelectedImage
        frmDesigner.TreeView1.Nodes.Item(udtSuperstructure.Key).EnsureVisible ' expand the tree branch
        'Update_Object udtSuperstructure.Key

        With colVehicle.Item(udtSuperstructure.Key) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
            .Key = frmDesigner.TreeView1.Nodes.Item(udtSuperstructure.Key).Key
            .Parent = frmDesigner.TreeView1.Nodes.Item(udtSuperstructure.Key).Parent.Key
            .Text = frmDesigner.TreeView1.Nodes.Item(udtSuperstructure.Key).Text
            .Image = frmDesigner.TreeView1.Nodes.Item(udtSuperstructure.Key).Image
            .SelectedImage = frmDesigner.TreeView1.Nodes.Item(udtSuperstructure.Key).SelectedImage
        End With

Case OpenMount ' OpenMount value is 10
        
        Get #1, , udtOpenMount
        i = Seek(1)
        'Create the object and add it to the collection and the tree
        Set objOpenMount = New clsOpenMount
        colVehicle.Add objOpenMount, udtOpenMount.Key  ' add the new object to the collection
        frmDesigner.TreeView1.Nodes.Add udtOpenMount.Parent, tvwChild, udtOpenMount.Key, udtOpenMount.Text, udtOpenMount.Image, udtOpenMount.SelectedImage
        frmDesigner.TreeView1.Nodes.Item(udtOpenMount.Key).EnsureVisible ' expand the tree branch
        'Update_Object udtOpenMount.Key

        With colVehicle.Item(udtOpenMount.Key) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
            .Key = frmDesigner.TreeView1.Nodes.Item(udtOpenMount.Key).Key
            .Parent = frmDesigner.TreeView1.Nodes.Item(udtOpenMount.Key).Parent.Key
            .Text = frmDesigner.TreeView1.Nodes.Item(udtOpenMount.Key).Text
            .Image = frmDesigner.TreeView1.Nodes.Item(udtOpenMount.Key).Image
            .SelectedImage = frmDesigner.TreeView1.Nodes.Item(udtOpenMount.Key).SelectedImage
        End With

Case Mast ' Mast value is 11
        
        Get #1, , udtMast
        i = Seek(1)
        'Create the object and add it to the collection and the tree
        Set objMast = New clsMast
        colVehicle.Add objMast, udtMast.Key  ' add the new object to the collection
        frmDesigner.TreeView1.Nodes.Add udtMast.Parent, tvwChild, udtMast.Key, udtMast.Text, udtMast.Image, udtMast.SelectedImage
        frmDesigner.TreeView1.Nodes.Item(udtMast.Key).EnsureVisible ' expand the tree branch
        'Update_Object udtMast.Key

        With colVehicle.Item(udtMast.Key) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
            .Key = frmDesigner.TreeView1.Nodes.Item(udtMast.Key).Key
            .Parent = frmDesigner.TreeView1.Nodes.Item(udtMast.Key).Parent.Key
            .Text = frmDesigner.TreeView1.Nodes.Item(udtMast.Key).Text
            .Image = frmDesigner.TreeView1.Nodes.Item(udtMast.Key).Image
            .SelectedImage = frmDesigner.TreeView1.Nodes.Item(udtMast.Key).SelectedImage
        End With

Case Pod ' Pod value is 12
        
        Get #1, , udtPod
        i = Seek(1)
        'Create the object and add it to the collection and the tree
        Set objPod = New clsPod
        colVehicle.Add objPod, udtPod.Key  ' add the new object to the collection
        frmDesigner.TreeView1.Nodes.Add udtPod.Parent, tvwChild, udtPod.Key, udtPod.Text, udtPod.Image, udtPod.SelectedImage
        frmDesigner.TreeView1.Nodes.Item(udtPod.Key).EnsureVisible ' expand the tree branch
        'Update_Object udtPod.Key

        With colVehicle.Item(udtPod.Key) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
            .Key = frmDesigner.TreeView1.Nodes.Item(udtPod.Key).Key
            .Parent = frmDesigner.TreeView1.Nodes.Item(udtPod.Key).Parent.Key
            .Text = frmDesigner.TreeView1.Nodes.Item(udtPod.Key).Text
            .Image = frmDesigner.TreeView1.Nodes.Item(udtPod.Key).Image
            .SelectedImage = frmDesigner.TreeView1.Nodes.Item(udtPod.Key).SelectedImage
        End With

Case Turret ' Turret value is 13
        
        Get #1, , udtTurret
        i = Seek(1)
        'Create the object and add it to the collection and the tree
        Set objTurret = New clsTurret
        colVehicle.Add objTurret, udtTurret.Key  ' add the new object to the collection
        frmDesigner.TreeView1.Nodes.Add udtTurret.Parent, tvwChild, udtTurret.Key, udtTurret.Text, udtTurret.Image, udtTurret.SelectedImage
        frmDesigner.TreeView1.Nodes.Item(udtTurret.Key).EnsureVisible ' expand the tree branch
        'Update_Object udtTurret.Key

        With colVehicle.Item(udtTurret.Key) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
            .Key = frmDesigner.TreeView1.Nodes.Item(udtTurret.Key).Key
            .Parent = frmDesigner.TreeView1.Nodes.Item(udtTurret.Key).Parent.Key
            .Text = frmDesigner.TreeView1.Nodes.Item(udtTurret.Key).Text
            .Image = frmDesigner.TreeView1.Nodes.Item(udtTurret.Key).Image
            .SelectedImage = frmDesigner.TreeView1.Nodes.Item(udtTurret.Key).SelectedImage
        End With

Case Popturret ' Popturret value is 14
        
        Get #1, , udtPopturret
        i = Seek(1)
        'Create the object and add it to the collection and the tree
        Set objPopturret = New clsPopTurret
        colVehicle.Add objPopturret, udtPopturret.Key  ' add the new object to the collection
        frmDesigner.TreeView1.Nodes.Add udtPopturret.Parent, tvwChild, udtPopturret.Key, udtPopturret.Text, udtPopturret.Image, udtPopturret.SelectedImage
        frmDesigner.TreeView1.Nodes.Item(udtPopturret.Key).EnsureVisible ' expand the tree branch
        'Update_Object udtPopturret.Key

        With colVehicle.Item(udtPopturret.Key) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
            .Key = frmDesigner.TreeView1.Nodes.Item(udtPopturret.Key).Key
            .Parent = frmDesigner.TreeView1.Nodes.Item(udtPopturret.Key).Parent.Key
            .Text = frmDesigner.TreeView1.Nodes.Item(udtPopturret.Key).Text
            .Image = frmDesigner.TreeView1.Nodes.Item(udtPopturret.Key).Image
            .SelectedImage = frmDesigner.TreeView1.Nodes.Item(udtPopturret.Key).SelectedImage
        End With

Case Gasbag ' Gasbag value is 15
        
        Get #1, , udtGasbag
        i = Seek(1)
        'Create the object and add it to the collection and the tree
        Set objGasbag = New clsGasbag
        colVehicle.Add objGasbag, udtGasbag.Key  ' add the new object to the collection
        frmDesigner.TreeView1.Nodes.Add udtGasbag.Parent, tvwChild, udtGasbag.Key, udtGasbag.Text, udtGasbag.Image, udtGasbag.SelectedImage
        frmDesigner.TreeView1.Nodes.Item(udtGasbag.Key).EnsureVisible ' expand the tree branch
        'Update_Object udtGasbag.Key

        With colVehicle.Item(udtGasbag.Key) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
            .Key = frmDesigner.TreeView1.Nodes.Item(udtGasbag.Key).Key
            .Parent = frmDesigner.TreeView1.Nodes.Item(udtGasbag.Key).Parent.Key
            .Text = frmDesigner.TreeView1.Nodes.Item(udtGasbag.Key).Text
            .Image = frmDesigner.TreeView1.Nodes.Item(udtGasbag.Key).Image
            .SelectedImage = frmDesigner.TreeView1.Nodes.Item(udtGasbag.Key).SelectedImage
        End With

Case Arm ' Arm value is 16
        
        Get #1, , udtArm
        i = Seek(1)
        'Create the object and add it to the collection and the tree
        Set objArm = New clsArm
        colVehicle.Add objArm, udtArm.Key  ' add the new object to the collection
        frmDesigner.TreeView1.Nodes.Add udtArm.Parent, tvwChild, udtArm.Key, udtArm.Text, udtArm.Image, udtArm.SelectedImage
        frmDesigner.TreeView1.Nodes.Item(udtArm.Key).EnsureVisible ' expand the tree branch
        'Update_Object udtArm.Key

        With colVehicle.Item(udtArm.Key) ' some of these values below are fixed and probably shouldn't need to include them all. That should speed it up
            .Key = frmDesigner.TreeView1.Nodes.Item(udtArm.Key).Key
            .Parent = frmDesigner.TreeView1.Nodes.Item(udtArm.Key).Parent.Key
            .Text = frmDesigner.TreeView1.Nodes.Item(udtArm.Key).Text
            .Image = frmDesigner.TreeView1.Nodes.Item(udtArm.Key).Image
            .SelectedImage = frmDesigner.TreeView1.Nodes.Item(udtArm.Key).SelectedImage
        End With

    End Select
Wend

Close #1

End Sub


Sub WriteRecord(TheFile As String)
' Writes the vehicle collection to a file specified by the user
    
    Dim iTmp As Integer
    Dim iIndex As Integer
    Dim iNext As Integer


' Open and delete any old data that exists in the file
Open TheFile For Output As #1
Close #1

' Open the file for binary access
Open TheFile For Binary As #1 ' need to replace this hardcoded path

MsgBox "There are " & colVehicle.Count & " objects in the collection" 'debug

    GetFirstParent 'Find a root node in the treeview
    'get the index of the root node that is at the top of the treeview
    iIndex = frmDesigner.TreeView1.Nodes(mnIndex).FirstSibling.Index
   
   iTmp = iIndex
            
' debug I really dont need a select case here since the first node is always the Body
 Select Case TypeName(colVehicle.Item(iIndex))
                            
         Case "clsBody"
                With colVehicle.Item(iIndex)
                      udtBody.Key = .Key
                      'udtBody.Parent = .Parent
                      udtBody.Image = .Image
                      udtBody.SelectedImage = .SelectedImage
                      udtBody.Text = .Text
                      'udtBody.Name = .Name
                End With
                
                Put #1, , Body     ' put the leading class identifier
                Put #1, , udtBody  'put the contents of the UDT "Body"

        End Select
        
    
    'If the Node has Children call the sub that writes the children
    If frmDesigner.TreeView1.Nodes(iIndex).Children > 0 Then
        WriteChild iIndex
    End If
    
    
    While iIndex <> frmDesigner.TreeView1.Nodes(iTmp).LastSibling.Index
    'loop through all the root nodes
         iNext = frmDesigner.TreeView1.Nodes(iIndex).Next.Index
            Select Case TypeName(colVehicle.Item(iNext))
            ' debug probably will only have subassemblies here.  Can't have components as siblings to the root node
            Case "clsArtilleryandGun"
          
              MsgBox udtArtilleryandGun.Parent
                With colVehicle.Item(iNext)
                      udtArtilleryandGun.Key = .Key
                      udtArtilleryandGun.Parent = .Parent
                      udtArtilleryandGun.Image = .Image
                      udtArtilleryandGun.SelectedImage = .SelectedImage
                      udtArtilleryandGun.Text = .Text
                      udtArtilleryandGun.Name = .Name
                End With
            
                Put #1, , ArtilleryandGun     ' put the leading class identifier
                Put #1, , udtArtilleryandGun  'put the contents of the UDT "Body"

            Case "clsWheel"  ' debug
                                    
                  With colVehicle.Item(iNext)
                       udtWheel.Key = .Key
                       udtWheel.Parent = .Parent
                       udtWheel.Image = .Image
                       udtWheel.SelectedImage = .SelectedImage
                       udtWheel.Text = .Text
                       udtWheel.Name = .Name
                       udtWheel.Number = .Number
                       udtWheel.ImprovedSuspension = .ImprovedSuspension
                       udtWheel.Wheelblades = .Wheelblades
                       udtWheel.SnowTires = .SnowTires
                       udtWheel.RacingTires = .RacingTires
                       udtWheel.PunctureResistant = .PunctureResistant
                       udtWheel.RetractLocation = .RetractLocation
                       udtWheel.ImprovedBreaks = .ImprovedBreaks
                       udtWheel.AllwheelSteering = .AllwheelSteering
                       udtWheel.Smartwheels = .Smartwheels
                   End With
                   Put #1, , Wheel     ' put the leading class identifier
                   Put #1, , udtWheel  'put the contents of the UDT "Wheel"
                   
        Case "clsSkid"
                                   
               With colVehicle.Item(iNext)
                      udtSkid.Key = .Key
                      udtSkid.Parent = .Parent
                      udtSkid.Image = .Image
                      udtSkid.SelectedImage = .SelectedImage
                      udtSkid.Text = .Text
                      udtSkid.Name = .Name
                End With
                Put #1, , Skid     ' put the leading class identifier
                Put #1, , udtSkid  'put the contents of the UDT "Body"
                
        Case "clsTrack"
                                   
               With colVehicle.Item(iNext)
                      udtTrack.Key = .Key
                      udtTrack.Parent = .Parent
                      udtTrack.Image = .Image
                      udtTrack.SelectedImage = .SelectedImage
                      udtTrack.Text = .Text
                      udtTrack.Name = .Name
                End With
                Put #1, , Track     ' put the leading class identifier
                Put #1, , udtTrack  'put the contents of the UDT "Body"
        
        Case "clsleg"
                                   
               With colVehicle.Item(iNext)
                      udtLeg.Key = .Key
                      udtLeg.Parent = .Parent
                      udtLeg.Image = .Image
                      udtLeg.SelectedImage = .SelectedImage
                      udtLeg.Text = .Text
                      udtLeg.Name = .Name
                End With
                Put #1, , Leg     ' put the leading class identifier
                Put #1, , udtLeg  'put the contents of the UDT "Body"
            
        Case "clsWing"
                                   
               With colVehicle.Item(iNext)
                      udtWing.Key = .Key
                      udtWing.Parent = .Parent
                      udtWing.Image = .Image
                      udtWing.SelectedImage = .SelectedImage
                      udtWing.Text = .Text
                      udtWing.Name = .Name
                End With
                Put #1, , Wing     ' put the leading class identifier
                Put #1, , udtWing  'put the contents of the UDT "Body"
            
        Case "clsRotar"
                                   
               With colVehicle.Item(iNext)
                      udtRotar.Key = .Key
                      udtRotar.Parent = .Parent
                      udtRotar.Image = .Image
                      udtRotar.SelectedImage = .SelectedImage
                      udtRotar.Text = .Text
                      udtRotar.Name = .Name
                End With
                Put #1, , Rotar     ' put the leading class identifier
                Put #1, , udtRotar  'put the contents of the UDT "Body"
            
        Case "clsHydrofoil"
                                   
               With colVehicle.Item(iNext)
                      udtHydrofoil.Key = .Key
                      udtHydrofoil.Parent = .Parent
                      udtHydrofoil.Image = .Image
                      udtHydrofoil.SelectedImage = .SelectedImage
                      udtHydrofoil.Text = .Text
                      udtHydrofoil.Name = .Name
                End With
                Put #1, , Hydrofoil     ' put the leading class identifier
                Put #1, , udtHydrofoil  'put the contents of the UDT "Body"
            
        Case "clsHovercraft"
                                   
               With colVehicle.Item(iNext)
                      udtHovercraft.Key = .Key
                      udtHovercraft.Parent = .Parent
                      udtHovercraft.Image = .Image
                      udtHovercraft.SelectedImage = .SelectedImage
                      udtHovercraft.Text = .Text
                      udtHovercraft.Name = .Name
                End With
                Put #1, , Hovercraft     ' put the leading class identifier
                Put #1, , udtHovercraft  'put the contents of the UDT "Body"
            
        Case "clsSuperstructure"
                                   
               With colVehicle.Item(iNext)
                      udtSuperstructure.Key = .Key
                      udtSuperstructure.Parent = .Parent
                      udtSuperstructure.Image = .Image
                      udtSuperstructure.SelectedImage = .SelectedImage
                      udtSuperstructure.Text = .Text
                      udtSuperstructure.Name = .Name
                End With
                Put #1, , Superstructure     ' put the leading class identifier
                Put #1, , udtSuperstructure  'put the contents of the UDT "Body"
            
        Case "clsOpenMount"
                                   
               With colVehicle.Item(iNext)
                      udtOpenMount.Key = .Key
                      udtOpenMount.Parent = .Parent
                      udtOpenMount.Image = .Image
                      udtOpenMount.SelectedImage = .SelectedImage
                      udtOpenMount.Text = .Text
                      udtOpenMount.Name = .Name
                End With
                Put #1, , OpenMount     ' put the leading class identifier
                Put #1, , udtOpenMount  'put the contents of the UDT "Body"
            
        Case "clsMast"
                                   
               With colVehicle.Item(iNext)
                      udtMast.Key = .Key
                      udtMast.Parent = .Parent
                      udtMast.Image = .Image
                      udtMast.SelectedImage = .SelectedImage
                      udtMast.Text = .Text
                      udtMast.Name = .Name
                End With
                Put #1, , Mast     ' put the leading class identifier
                Put #1, , udtMast  'put the contents of the UDT "Body"
            
        Case "clsPod"
                                   
               With colVehicle.Item(iNext)
                      udtPod.Key = .Key
                      udtPod.Parent = .Parent
                      udtPod.Image = .Image
                      udtPod.SelectedImage = .SelectedImage
                      udtPod.Text = .Text
                      udtPod.Name = .Name
                End With
                Put #1, , Pod     ' put the leading class identifier
                Put #1, , udtPod  'put the contents of the UDT "Body"
            
        Case "clsTurret"
                                   
               With colVehicle.Item(iNext)
                      udtTurret.Key = .Key
                      udtTurret.Parent = .Parent
                      udtTurret.Image = .Image
                      udtTurret.SelectedImage = .SelectedImage
                      udtTurret.Text = .Text
                      udtTurret.Name = .Name
                End With
                Put #1, , Turret     ' put the leading class identifier
                Put #1, , udtTurret  'put the contents of the UDT "Body"
                
        Case "clsPopturret"
                                   
               With colVehicle.Item(iNext)
                      udtPopturret.Key = .Key
                      udtPopturret.Parent = .Parent
                      udtPopturret.Image = .Image
                      udtPopturret.SelectedImage = .SelectedImage
                      udtPopturret.Text = .Text
                      udtPopturret.Name = .Name
                End With
                Put #1, , Popturret     ' put the leading class identifier
                Put #1, , udtPopturret  'put the contents of the UDT "Body"
            
        Case "clsGasbag"
                                   
               With colVehicle.Item(iNext)
                      udtGasbag.Key = .Key
                      udtGasbag.Parent = .Parent
                      udtGasbag.Image = .Image
                      udtGasbag.SelectedImage = .SelectedImage
                      udtGasbag.Text = .Text
                      udtGasbag.Name = .Name
                End With
                Put #1, , Gasbag     ' put the leading class identifier
                Put #1, , udtGasbag  'put the contents of the UDT "Body"
            
        Case "clsArm"
                                   
               With colVehicle.Item(iNext)
                      udtArm.Key = .Key
                      udtArm.Parent = .Parent
                      udtArm.Image = .Image
                      udtArm.SelectedImage = .SelectedImage
                      udtArm.Text = .Text
                      udtArm.Name = .Name
                End With
                Put #1, , Arm     ' put the leading class identifier
                Put #1, , udtArm  'put the contents of the UDT "Body"
         End Select
            
        'If the Node has Children call the sub that writes the children
        If frmDesigner.TreeView1.Nodes(iIndex).Next.Children > 0 Then
            WriteChild frmDesigner.TreeView1.Nodes(iIndex).Next.Index
        End If
        
        ' Move to the Next root Node
        iIndex = frmDesigner.TreeView1.Nodes(iIndex).Next.Index
    Wend
' Close the file
Close #1
End Sub

Private Sub WriteChild(ByVal iNodeIndex As Integer)
' Write the child nodes to the table. This sub uses recursion
' to loop through the child nodes. It receives the Index of
' the node that has the children
    Dim i As Integer
    Dim iTempIndex As Integer
    iTempIndex = frmDesigner.TreeView1.Nodes(iNodeIndex).Child.FirstSibling.Index
    'Loop through all a Parents Child Nodes
    For i = 1 To frmDesigner.TreeView1.Nodes(iNodeIndex).Children
        
            Select Case TypeName(colVehicle.Item(iTempIndex))
                            
            Case "clsArtilleryandGun"
            
                With colVehicle.Item(iTempIndex)
                      udtArtilleryandGun.Key = .Key
                      udtArtilleryandGun.Parent = .Parent
                      udtArtilleryandGun.Image = .Image
                      udtArtilleryandGun.SelectedImage = .SelectedImage
                      udtArtilleryandGun.Text = .Text
                      udtArtilleryandGun.Name = .Name
                End With
             
                Put #1, , ArtilleryandGun     ' put the leading class identifier
                Put #1, , udtArtilleryandGun  'put the contents of the UDT "Body"
            
            Case "clsWheel"  ' debug
                                    
                  With colVehicle.Item(iTempIndex)
                       udtWheel.Key = .Key
                       udtWheel.Parent = .Parent
                       udtWheel.Image = .Image
                       udtWheel.SelectedImage = .SelectedImage
                       udtWheel.Text = .Text
                       udtWheel.Name = .Name
                       udtWheel.Number = .Number
                       udtWheel.ImprovedSuspension = .ImprovedSuspension
                       udtWheel.Wheelblades = .Wheelblades
                       udtWheel.SnowTires = .SnowTires
                       udtWheel.RacingTires = .RacingTires
                       udtWheel.PunctureResistant = .PunctureResistant
                       udtWheel.RetractLocation = .RetractLocation
                       udtWheel.ImprovedBreaks = .ImprovedBreaks
                       udtWheel.AllwheelSteering = .AllwheelSteering
                       udtWheel.Smartwheels = .Smartwheels
                   End With
                   Put #1, , Wheel     ' put the leading class identifier
                   Put #1, , udtWheel  'put the contents of the UDT "Wheel"
                   
                   
                  
         Case "clsBody"
                                   
               With colVehicle.Item(iTempIndex)
                      udtBody.Key = .Key
                      udtBody.Parent = .Parent
                      udtBody.Image = .Image
                      udtBody.SelectedImage = .SelectedImage
                      udtBody.Text = .Text
                      udtBody.Name = .Name
                End With
                Put #1, , Body     ' put the leading class identifier
                Put #1, , udtBody  'put the contents of the UDT "Body"
        
        Case "clsSkid"
                                   
               With colVehicle.Item(iTempIndex)
                      udtSkid.Key = .Key
                      udtSkid.Parent = .Parent
                      udtSkid.Image = .Image
                      udtSkid.SelectedImage = .SelectedImage
                      udtSkid.Text = .Text
                      'udtSkid.Name = .Name
                End With
                Put #1, , Skid     ' put the leading class identifier
                Put #1, , udtSkid  'put the contents of the UDT "Body"
                
        Case "clsTrack"
                                   
               With colVehicle.Item(iTempIndex)
                      udtTrack.Key = .Key
                      udtTrack.Parent = .Parent
                      udtTrack.Image = .Image
                      udtTrack.SelectedImage = .SelectedImage
                      udtTrack.Text = .Text
                      udtTrack.Name = .Name
                End With
                Put #1, , Track     ' put the leading class identifier
                Put #1, , udtTrack  'put the contents of the UDT "Body"
        
        Case "clsleg"
                                   
               With colVehicle.Item(iTempIndex)
                      udtLeg.Key = .Key
                      udtLeg.Parent = .Parent
                      udtLeg.Image = .Image
                      udtLeg.SelectedImage = .SelectedImage
                      udtLeg.Text = .Text
                      udtLeg.Name = .Name
                End With
                Put #1, , Leg     ' put the leading class identifier
                Put #1, , udtLeg  'put the contents of the UDT "Body"
            
        Case "clsWing"
                                   
               With colVehicle.Item(iTempIndex)
                      udtWing.Key = .Key
                      udtWing.Parent = .Parent
                      udtWing.Image = .Image
                      udtWing.SelectedImage = .SelectedImage
                      udtWing.Text = .Text
                      udtWing.Name = .Name
                End With
                Put #1, , Wing     ' put the leading class identifier
                Put #1, , udtWing  'put the contents of the UDT "Body"
            
        Case "clsRotar"
                                   
               With colVehicle.Item(iTempIndex)
                      udtRotar.Key = .Key
                      udtRotar.Parent = .Parent
                      udtRotar.Image = .Image
                      udtRotar.SelectedImage = .SelectedImage
                      udtRotar.Text = .Text
                      udtRotar.Name = .Name
                End With
                Put #1, , Rotar     ' put the leading class identifier
                Put #1, , udtRotar  'put the contents of the UDT "Body"
            
        Case "clsHydrofoil"
                                   
               With colVehicle.Item(iTempIndex)
                      udtHydrofoil.Key = .Key
                      udtHydrofoil.Parent = .Parent
                      udtHydrofoil.Image = .Image
                      udtHydrofoil.SelectedImage = .SelectedImage
                      udtHydrofoil.Text = .Text
                      udtHydrofoil.Name = .Name
                End With
                Put #1, , Hydrofoil     ' put the leading class identifier
                Put #1, , udtHydrofoil  'put the contents of the UDT "Body"
            
        Case "clsHovercraft"
                                   
               With colVehicle.Item(iTempIndex)
                      udtHovercraft.Key = .Key
                      udtHovercraft.Parent = .Parent
                      udtHovercraft.Image = .Image
                      udtHovercraft.SelectedImage = .SelectedImage
                      udtHovercraft.Text = .Text
                      udtHovercraft.Name = .Name
                End With
                Put #1, , Hovercraft     ' put the leading class identifier
                Put #1, , udtHovercraft  'put the contents of the UDT "Body"
            
        Case "clsSuperstructure"
                                   
               With colVehicle.Item(iTempIndex)
                      udtSuperstructure.Key = .Key
                      udtSuperstructure.Parent = .Parent
                      udtSuperstructure.Image = .Image
                      udtSuperstructure.SelectedImage = .SelectedImage
                      udtSuperstructure.Text = .Text
                      udtSuperstructure.Name = .Name
                End With
                Put #1, , Superstructure     ' put the leading class identifier
                Put #1, , udtSuperstructure  'put the contents of the UDT "Body"
            
        Case "clsOpenMount"
                                   
               With colVehicle.Item(iTempIndex)
                      udtOpenMount.Key = .Key
                      udtOpenMount.Parent = .Parent
                      udtOpenMount.Image = .Image
                      udtOpenMount.SelectedImage = .SelectedImage
                      udtOpenMount.Text = .Text
                      'udtOpenMount.Name = .Name
                End With
                Put #1, , OpenMount     ' put the leading class identifier
                Put #1, , udtOpenMount  'put the contents of the UDT "Body"
            
        Case "clsMast"
                                   
               With colVehicle.Item(iTempIndex)
                      udtMast.Key = .Key
                      udtMast.Parent = .Parent
                      udtMast.Image = .Image
                      udtMast.SelectedImage = .SelectedImage
                      udtMast.Text = .Text
                      'udtMast.Name = .Name
                End With
                Put #1, , Mast     ' put the leading class identifier
                Put #1, , udtMast  'put the contents of the UDT "Body"
            
        Case "clsPod"
                                   
               With colVehicle.Item(iTempIndex)
                      udtPod.Key = .Key
                      udtPod.Parent = .Parent
                      udtPod.Image = .Image
                      udtPod.SelectedImage = .SelectedImage
                      udtPod.Text = .Text
                      udtPod.Name = .Name
                End With
                Put #1, , Pod     ' put the leading class identifier
                Put #1, , udtPod  'put the contents of the UDT "Body"
            
        Case "clsTurret"
                                   
               With colVehicle.Item(iTempIndex)
                      udtTurret.Key = .Key
                      udtTurret.Parent = .Parent
                      udtTurret.Image = .Image
                      udtTurret.SelectedImage = .SelectedImage
                      udtTurret.Text = .Text
                      udtTurret.Name = .Name
                End With
                Put #1, , Turret     ' put the leading class identifier
                Put #1, , udtTurret  'put the contents of the UDT "Body"
                
        Case "clsPopturret"
                                   
               With colVehicle.Item(iTempIndex)
                      udtPopturret.Key = .Key
                      udtPopturret.Parent = .Parent
                      udtPopturret.Image = .Image
                      udtPopturret.SelectedImage = .SelectedImage
                      udtPopturret.Text = .Text
                      udtPopturret.Name = .Name
                End With
                Put #1, , Popturret     ' put the leading class identifier
                Put #1, , udtPopturret  'put the contents of the UDT "Body"
            
        Case "clsGasbag"
                                   
               With colVehicle.Item(iTempIndex)
                      udtGasbag.Key = .Key
                      udtGasbag.Parent = .Parent
                      udtGasbag.Image = .Image
                      udtGasbag.SelectedImage = .SelectedImage
                      udtGasbag.Text = .Text
                      udtGasbag.Name = .Name
                End With
                Put #1, , Gasbag     ' put the leading class identifier
                Put #1, , udtGasbag  'put the contents of the UDT "Body"
            
        Case "clsArm"
                                   
               With colVehicle.Item(iTempIndex)
                      udtArm.Key = .Key
                      udtArm.Parent = .Parent
                      udtArm.Image = .Image
                      udtArm.SelectedImage = .SelectedImage
                      udtArm.Text = .Text
                      udtArm.Name = .Name
                End With
                Put #1, , Arm     ' put the leading class identifier
                Put #1, , udtArm  'put the contents of the UDT "Body"
         End Select
        
  ' If the Node we are on has a child call the Sub again
    If frmDesigner.TreeView1.Nodes(iTempIndex).Children > 0 Then
       WriteChild (iTempIndex)
    End If
    
  ' If we are not on the last child move to the next child Node
    If i <> frmDesigner.TreeView1.Nodes(iNodeIndex).Children Then
       iTempIndex = frmDesigner.TreeView1.Nodes(iTempIndex).Next.Index
    End If
  Next i
End Sub


Sub ReadINI()
' MAKE SURE that i destroy the UDT after reading. Dont need to
' Keep this in memory the whole time the program is running.  The
' .initial directory can be made a regular variable... doesnt need
' to be in a UDT

'NOTE: I should probably clean up both ReadINI and WriteINI to not include
' any UDT's and only read the direct properties of objects and put them
' using sequential access.  WIll be easy to do

' Make sure the dat (aka INI) file exists
If Dir(App.Path + "\designer.dat") = "designer.dat" Then ' debug  this line needs to change because its saving the INI in the wrong place
    
    Open "designer.dat" For Random As #2
    Get #2, 1, udtSettings
    Close 2
    
    ' If the INI does not exist, create the file
Else
    MsgBox "the ini does not exist" ' debug
   
    ' Since there was no designer.dat file, fill udtSettings with default settings
    With udtSettings
        .InitialDir = App.Path + "\vehicles"
    End With
    
End If

End Sub


Sub WriteINI()

' switch to the apps directory
Dir App.Path
' open the ini
Open "designer.dat" For Random As #2

' update the form locations info
With udtSettings
End With

' save the udtSettings data and close the file
Put #2, 1, udtSettings
Close 2

End Sub

