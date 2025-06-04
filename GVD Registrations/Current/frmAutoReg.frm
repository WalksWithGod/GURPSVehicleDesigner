VERSION 5.00
Begin VB.Form frmAutoReg 
   Caption         =   "Form1"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14775
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   14775
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   23
      Text            =   "D:\visualb\projects\GVD Registrations\Current\[GVD] Thanks for registering!.txt"
      Top             =   360
      Width           =   6135
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7680
      Width           =   2820
   End
   Begin VB.TextBox txtCustomerID 
      Height          =   285
      Left            =   1560
      TabIndex        =   21
      Text            =   "txtCustomerID"
      Top             =   3720
      Width           =   2655
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Index           =   3
      Left            =   3240
      TabIndex        =   19
      Text            =   "txtName(3)"
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   18
      Text            =   "txtName(2)"
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   17
      Text            =   "txtName(1)"
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   16
      Text            =   "txtName(0)"
      Top             =   2640
      Width           =   2535
   End
   Begin VB.CommandButton cmdSendEmail 
      Caption         =   "Add Record + Send Email"
      Height          =   495
      Left            =   11640
      TabIndex        =   11
      Top             =   7680
      Width           =   1335
   End
   Begin VB.TextBox txtRegCode 
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Text            =   "txtRegCode"
      Top             =   4080
      Width           =   2775
   End
   Begin VB.TextBox txtOrder 
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   8
      Text            =   "NOW"
      Top             =   2040
      Width           =   2775
   End
   Begin VB.TextBox txtOrder 
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   7
      Text            =   "1"
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox txtOrder 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   6
      Text            =   "1"
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "D:\visualb\projects\GVD Registrations\Current\GVD 2000.mdb"
      Top             =   840
      Width           =   4575
   End
   Begin VB.TextBox txtRegSoft 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7065
      Left            =   6600
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Text            =   "frmAutoReg.frx":0000
      Top             =   360
      Width           =   5985
   End
   Begin VB.Label Label9 
      Caption         =   "Email Template:"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Customer ID"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "User Middle"
      Height          =   255
      Left            =   1680
      TabIndex        =   15
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "User Last"
      Height          =   255
      Left            =   3480
      TabIndex        =   14
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "User First"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "User Full Name"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   4560
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4560
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label Label3 
      Caption         =   "Reg Code"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Order Date"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Product ID"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblEmployeeID 
      Caption         =   "Employee ID"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblDataSrc 
      Caption         =   "Data Source:"
      Height          =   255
      Left            =   90
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmAutoReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sRegSoftEmailTxt As String

Dim m_sBuffer() As String

Private Type OrderInfo

    ORDERID As String
    TRACKINGID As String
    ORDERDATE As String '
    PRODUCTNAME As String
    PRODUCTID As String
    ORDERTYPE As String
    QUANTITY As Long
    REGTOTAL As Currency
End Type

Private Type CustomerInfo
    Name As String
    COMPANY As String
    Email As String
    ADDRESS1 As String
    ADDRESS2 As String
    ADDRESS3 As String
    CITY As String
    STATE As String
    COUNTRY As String
    COUNTRYCODE As String
    ZIP As String
    PHONE As String
End Type

Dim m_uOrderInfo As OrderInfo
Dim m_uCustomerInfo As CustomerInfo


' Keywords

Const ORDERID = "Order ID"
Const TRACKINGID = "Tracking ID"
Const REGDATE = "Registration Date"
Const PRODUCTNAME = "Product Name"
Const PRODUCTID = "Product ID"
Const ORDERTYPE = "Order Type"
Const QUANTITY = "Quantity"
Const REGTOTAL = "Registration Total"


Const CUSTOMERNAME = "Name"
Const COMPANY = "Company"
Const Email = "Email"
Const ADDRESS1 = "Address1"
Const ADDRESS2 = "Address2"
Const ADDRESS3 = "Address3"
Const CITY = "City"
Const STATE = "State"
Const COUNTRY = "Country"
Const COUNTRYCODE = "Country Code"
Const ZIP = "Zip"
Const PHONE = "Phone"


Private Sub Form_Load()
    txtOrder(2) = Now
End Sub

Private Sub txtRegSoft_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        txtRegSoft.OLEDrag
    End If
End Sub

Private Sub txtRegSoft_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim hFile As Long
    
    Dim lngFileLen As Long
    Dim lngCount As Long
    
    hFile = FreeFile
    
    
    sRegSoftEmailTxt = Data.Files(1)
    lngFileLen = FileLen(sRegSoftEmailTxt)
    
    txtRegSoft.Text = ""
    Open sRegSoftEmailTxt For Input As #hFile
    
   Do While Not EOF(hFile)
        
        ReDim Preserve m_sBuffer(lngCount)
        Line Input #hFile, m_sBuffer(lngCount)
        txtRegSoft.Text = txtRegSoft.Text & vbNewLine & m_sBuffer(lngCount)
        lngCount = lngCount + 1
   Loop

   Close #hFile
    
   Call ParseRegInfo
   
End Sub

Private Sub ParseRegInfo()
      
    
    With m_uOrderInfo
        .ORDERID = GetValue(ORDERID)
        .TRACKINGID = GetValue(TRACKINGID)
       ' .ORDERDATE = GetValue(ORDERDATE)
        .PRODUCTNAME = GetValue(PRODUCTNAME)
        .PRODUCTID = GetValue(PRODUCTID)
        .ORDERTYPE = GetValue(ORDERTYPE)
        .QUANTITY = Val(GetValue(QUANTITY))
        .REGTOTAL = CCur(GetValue(REGTOTAL))
    End With
    
    With m_uCustomerInfo
        .Name = GetValue(CUSTOMERNAME)
        .COMPANY = GetValue(COMPANY)
        .Email = GetValue(Email)
        .ADDRESS1 = GetValue(ADDRESS1)
        .ADDRESS2 = GetValue(ADDRESS2)
        .ADDRESS3 = GetValue(ADDRESS3)
        .CITY = GetValue(CITY)
        .STATE = GetValue(STATE)
        .COUNTRY = GetValue(COUNTRY)
        .COUNTRYCODE = GetValue(COUNTRYCODE)
        .ZIP = GetValue(ZIP)
        .PHONE = GetValue(PHONE)
    End With

End Sub

Private Function GetValue(sSearch As String) As String
    Dim s() As String
    
    Dim i As Long
    Dim lRet As Long
    
    Const START_OFFSET = 1
    Const DELIM = ":"
    
    For i = 0 To UBound(m_sBuffer)
        lRet = InStr(START_OFFSET, m_sBuffer(i), sSearch)
        If lRet Then
            s = Split(m_sBuffer(i), DELIM)
            GetValue = Trim(s(1))
            Exit For
        End If
    Next
End Function
