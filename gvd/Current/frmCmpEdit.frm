VERSION 5.00
Object = "{FF047D84-C3F1-11D2-877E-0040055C08D9}#1.0#0"; "TreeX.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCmpEdit 
   Caption         =   "Component Editor - untitled.xml"
   ClientHeight    =   7200
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9600
   Icon            =   "frmCmpEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   11245
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "basics"
      TabPicture(0)   =   "frmCmpEdit.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(3)=   "Frame4"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "properties"
      TabPicture(1)   =   "frmCmpEdit.frx":045E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "treeProperties"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lstProperties"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command1(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command2(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Command2(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "rules"
      TabPicture(2)   =   "frmCmpEdit.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label4"
      Tab(2).Control(1)=   "Label5"
      Tab(2).Control(2)=   "MSHFlexGrid1"
      Tab(2).ControlCount=   3
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   4335
         Left            =   -71160
         TabIndex        =   34
         Top             =   1320
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   7646
         _Version        =   393216
         Rows            =   4
         Cols            =   4
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Move Down"
         Height          =   495
         Index           =   1
         Left            =   6000
         TabIndex        =   32
         Top             =   5520
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Move Up"
         Height          =   495
         Index           =   0
         Left            =   4800
         TabIndex        =   31
         Top             =   5520
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "New Modifier"
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   30
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         Caption         =   "Notes"
         Height          =   2895
         Left            =   -70560
         TabIndex        =   28
         Top             =   3240
         Width           =   4695
         Begin VB.TextBox txtNotes 
            Height          =   2325
            Left            =   240
            TabIndex        =   29
            Text            =   "txtNotes"
            Top             =   360
            Width           =   4095
         End
      End
      Begin VB.ListBox lstProperties 
         Height          =   4545
         Left            =   4680
         TabIndex        =   27
         Top             =   480
         Width           =   4335
      End
      Begin VB.Frame Frame3 
         Caption         =   "Component"
         Height          =   2895
         Left            =   -74640
         TabIndex        =   11
         Top             =   3240
         Width           =   3975
         Begin VB.Label Label2 
            Caption         =   "Deleteable?"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   16
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Icon Path:"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   15
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Name:"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   14
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Path:"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   13
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Classname:"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   12
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "File Info"
         Height          =   2415
         Left            =   -70440
         TabIndex        =   10
         Top             =   720
         Width           =   4335
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   2
            Left            =   3720
            TabIndex        =   26
            Text            =   "Text2"
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   1
            Left            =   3240
            TabIndex        =   25
            Text            =   "Text2"
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   5
            Left            =   1080
            TabIndex        =   22
            Text            =   "Text2"
            Top             =   1440
            Width           =   3135
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   4
            Left            =   1080
            TabIndex        =   21
            Text            =   "Text2"
            Top             =   1080
            Width           =   3135
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   20
            Text            =   "Text2"
            Top             =   720
            Width           =   3135
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   0
            Left            =   2760
            TabIndex        =   18
            Text            =   "Text2"
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "Modified:"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   24
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Created:"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   23
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "GUID:"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   19
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label3 
            Caption         =   "Version (Major/Minor/Revision)"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   17
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Author Info"
         Height          =   2415
         Left            =   -74640
         TabIndex        =   3
         Top             =   720
         Width           =   3975
         Begin VB.TextBox txtAuthor 
            Height          =   285
            Index           =   5
            Left            =   2640
            TabIndex        =   41
            Text            =   "Text1"
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtAuthor 
            Height          =   285
            Index           =   4
            Left            =   840
            TabIndex        =   40
            Text            =   "Text1"
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtAuthor 
            Height          =   285
            Index           =   3
            Left            =   2640
            TabIndex        =   39
            Text            =   "Text1"
            Top             =   480
            Width           =   1095
         End
         Begin VB.TextBox txtAuthor 
            Height          =   285
            Index           =   2
            Left            =   840
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   1920
            Width           =   2895
         End
         Begin VB.TextBox txtAuthor 
            Height          =   285
            Index           =   1
            Left            =   840
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   1440
            Width           =   2895
         End
         Begin VB.TextBox txtAuthor 
            Height          =   285
            Index           =   0
            Left            =   840
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Nick"
            Height          =   255
            Index           =   5
            Left            =   2040
            TabIndex        =   38
            Top             =   840
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Middle"
            Height          =   255
            Index           =   4
            Left            =   2040
            TabIndex        =   37
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Last"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   36
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Email:"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   6
            Top             =   1920
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "URL: "
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   5
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "First"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   4
            Top             =   480
            Width           =   495
         End
      End
      Begin TreeXLibCtl.TreeX treeProperties 
         Height          =   4695
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   4095
         _cx             =   1368923191
         _cy             =   1368924249
         BorderStyle     =   5
         BackColor       =   16777215
         ForeColor       =   0
         PicturePosition =   17
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         MousePointer    =   0
         AutoHScroll     =   -1  'True
         AutoVScroll     =   -1  'True
      End
      Begin VB.Label Label5 
         Caption         =   $"frmCmpEdit.frx":0496
         Height          =   3375
         Left            =   -74520
         TabIndex        =   35
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   $"frmCmpEdit.frx":05E0
         Height          =   855
         Left            =   -74400
         TabIndex        =   33
         Top             =   600
         Width           =   7935
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   6720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCmpEdit.frx":06D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCmpEdit.frx":07EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCmpEdit.frx":08FD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
   End
End
Attribute VB_Name = "frmCmpEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const DATA_PATH = "\Data\props_oldv.mdb" ' "\Data\properties.mdb"

Private m_oRoot As cBase

'rules ?
'location()
'designcheck()
'bitflags?

Dim m_oRS As ADODB.Recordset
Dim m_oDBConnection As ADODB.Connection
Dim m_oCommand As ADODB.Command

Dim WithEvents m_oObjStore As ObjectStore
Attribute m_oObjStore.VB_VarHelpID = -1


Private Sub Form_Load()
    Set m_oObjStore = New ObjectStore
    
    'connect to our DB
    Set m_oDBConnection = New ADODB.Connection
    
    m_oDBConnection.Open "Provider= Microsoft.Jet.OLEDB.3.51;" & "Data Source=" & App.Path & DATA_PATH
    
    
    
   ' set basic settings for the recordset
    Set m_oRS = New ADODB.Recordset
    m_oRS.CursorType = adOpenStatic
    m_oRS.CursorLocation = adUseClient
    m_oRS.LockType = adLockPessimistic    'This guarantees that a record that is being edited can be saved.
    m_oRS.ActiveConnection = m_oDBConnection    'The record set needs to know what connection to use.
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_oRoot = Nothing
    Set m_oObjStore = Nothing
    
    Set m_oRS = Nothing
    Set m_oDBConnection = Nothing
    Set m_oCommand = Nothing
End Sub

Private Sub m_oObjStore_RequestObject(ByVal className As String, ByVal DefPath As String, ByVal GUID As String, newObject As PersistenceManager.cIPersist)
    Set newObject = CreateObject(className)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Const BUTTON_OPEN = 2
    Const BUTTON_NEW = 1
    Const BUTTON_SAVE = 3
    
    Select Case Button.Index
    
        Case BUTTON_OPEN
            Dim oDlg As clsCmdlg
            Dim sFileName As String
            Dim oDisplay As cIDisplay
            Dim oProperties As cpropertyitem
            
            Set oDlg = New clsCmdlg
            If oDlg.ShowOpen(Me.hWnd) Then
                ' verify its a good xml file
                sFileName = oDlg.cFileName(1)
                
                Set m_oRoot = New cBase
                Set oDisplay = m_oRoot
                
                m_oObjStore.Deserialize sFileName, 0, XML_NODE_OBJECT, m_oRoot
                    
                ' todo: verify that deserialization succeeded?
                
                setFormCaption sFileName
                
                ' clear the tree and property list
                treeProperties.RemoveAllItems
                lstProperties.Clear
                
                ' retreive and display properties that are in the saved file (since they exist in the file, these are "visible" properties)
                Set oProperties = oDisplay.getFirstPropertyItem
                Do While Not oProperties Is Nothing
                    
                    lstProperties.AddItem oProperties.Caption
                    
                    Set oProperties = oDisplay.getNextPropertyItem
                Loop
                
                
                Dim sClassName As String
                Dim lngClassID As Long
                
                ' grab the classname.  We use that with our interfaces.mdb to find the classname in the classes table.
                sClassName = m_oRoot.className
                
                m_oRS.Source = "SELECT * FROM classes WHERE className = '" & sClassName & "'"
                m_oRS.Open
                lngClassID = m_oRS("id").Value
                
                '   retrieve the list of interfaces which this class supports.
                Dim lngInterfaces() As Long
                Dim strInterfaces() As String
                Dim lngNumInterfaces As Long
                Dim i As Long
                
                m_oRS.Close
                m_oRS.Source = "SELECT * FROM interfaces WHERE idClass = " & lngClassID
                m_oRS.Open
                
                lngNumInterfaces = m_oRS.RecordCount
                ReDim lngInterfaces(0 To lngNumInterfaces - 1)
                ReDim strInterfaces(0 To lngNumInterfaces - 1)
                m_oRS.MoveFirst
                For i = 0 To lngNumInterfaces - 1
                    lngInterfaces(i) = m_oRS("id").Value
                    strInterfaces(i) = m_oRS("name").Value
                    m_oRS.MoveNext
                Next
                
        
                '   for each supported interface retreive full list of properties
                Dim lngNumProperties As Long
                Dim lngProperties() As Long
                Dim j As Long
                Dim hParent As Long
                Dim lngCheckValue As Long
                
                Const DOWN = 1
                Const SIDE = 0
                
                For i = 0 To lngNumInterfaces - 1
                    m_oRS.Close
                    m_oRS.Source = "SELECT * FROM properties WHERE idInterface = " & lngInterfaces(i)
                    m_oRS.Open
                    
                    lngNumProperties = m_oRS.RecordCount
                    
                    m_oRS.MoveFirst
                    hParent = treeProperties.AddItem(strInterfaces(i), SIDE)
                   
                    For j = 0 To lngNumProperties - 1
                        ' add the property to the tree under the correct interface
                        treeProperties.AddItemCheck m_oRS("caption").Value, lngCheckValue, hParent, DOWN
                        m_oRS.MoveNext
                    Next
                    treeProperties.ExpandItem(hParent) = True
                Next
                m_oRS.Close
            End If
            
            Set oDlg = Nothing
            Set oDisplay = Nothing
            
            Call updateDisplay
        
        Case BUTTON_NEW
            
            
        Case BUTTON_SAVE
        
        
    End Select
End Sub

Private Sub setFormCaption(ByRef sFileName As String)
    If sFileName = "" Then sFileName = "untitled"
    frmCmpEdit.Caption = "Component Editor - " & sFileName
End Sub
Private Sub updateDisplay()

End Sub

Private Function loadNodes(oRoot As MSXML2.IXMLDOMNode) As Long
    On Error GoTo error
        
'    Set m_oAuthor = oRoot.selectSingleNode("author")
'    ' get the sub node values we need
'     txtAuthor(2) = m_oAuthor.selectSingleNode("string[@name='url']").nodeTypedValue
'
'    loadNodes = True
'    Exit Function
error:
    loadNodes = False
End Function

Private Function loadDef(sFile As String) As Boolean
    
    Dim oXDef As cXML
    Dim sGuid As String
    Dim oXNode As MSXML2.IXMLDOMNode
    '--
    
    'Now we need to open the XML Defintion file
    Set oXDef = New cXML
    
    If oXDef.OpenFromFile(sFile, True) Then
        Set m_oRoot = oXDef.GetRootNode
        Set oXDef = Nothing
        loadDef = True
    End If
    Exit Function
err:
        loadDef = False
       Debug.Print "modListViewHelper:ReadComponentXML -- Error #" & err.Number & " " & err.Description
End Function
