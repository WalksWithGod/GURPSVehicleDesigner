VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRegNum 
      Height          =   405
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2505
   End
   Begin VB.TextBox txtRegID 
      Height          =   285
      Left            =   1740
      TabIndex        =   2
      Top             =   120
      Width           =   405
   End
   Begin VB.TextBox txtRegName 
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   525
      Left            =   3480
      TabIndex        =   0
      Top             =   300
      Width           =   705
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gsRegName() As Byte
Dim gsRegNum() As Byte
Dim gsRegID As Long

Private Sub Command1_Click()
Dim i As Long

    gsRegID = Val(txtRegID) 'this is the "userID" of the reg'd user

    For i = 1 To Len(txtRegName)
        ReDim Preserve gsRegName(i)
        
        gsRegName(i) = Asc(Mid(txtRegName.Text, i, 1))
    Next
    For i = 1 To Len(txtRegNum)
        ReDim Preserve gsRegNum(i)
        gsRegNum(i) = Asc(Mid(txtRegNum.Text, i, 1))
    Next
    
    ChopCheck
End Sub

Public Function ChopCheck() As Byte()
Dim tempbyte() As Byte
Dim i As Long
Dim j As Single
Dim sTName As String
Dim lngtotal As Single
Dim sRegNumber As String

ReDim tempbyte(1)
tempbyte(1) = 116
'//one of the local reg number checkers.  There will be several of these so
' so that a hacker will have to do some serious code hacking to disable all
' of them

'here's the reg key formula
'1- the user's reg name and key are accepted into a byte array with each
'   letter being actually the ascii code for that letter.  Total them up
For i = 1 To UBound(gsRegName)
    lngtotal = lngtotal + gsRegName(i)
    'at the same time total the ascii value for every even valued ascii code
    If gsRegName(i) Mod 2 = 0 Then
        lngtotal = lngtotal + gsRegName(i)
    End If
Next
'2 - the RegID is actually just a modifier to prevent two people having the same
'    name winding up with the same ID.  This ID is unique and alone can be used
'   to identify a user.  Multiply this to the total
lngtotal = lngtotal * gsRegID
'3- take the ascii value of the typename of the Body and multiply that to it
sTName = "clsBody"
For i = 1 To Len(sTName)
    lngtotal = lngtotal * Asc(Mid(sTName, i, 1))
Next
'6- take a random seed to generate the seeded random number and multiply that
Rnd -1
Randomize 9921988
lngtotal = lngtotal * Rnd()
'8- return this as a byte array that we can compare with our current one
'how do we split this up into seperate bytes? well we know our ascii values
'must be between 48-57, 65-90 and 97-122
'well, we can generate a random reg code based on each number in the string
'representation using the random seed of each number
For i = 1 To Len(Str(lngtotal))
    j = Rnd()
    If j <= 0.33 Then
        ReDim Preserve tempbyte(i)
        Rnd -1
        Randomize Asc(Mid(Str(lngtotal), i, 1))
        tempbyte(i) = Int((57 - 48 + 1) * Rnd + 48)
        sRegNumber = sRegNumber & Chr(tempbyte(i))
    ElseIf j <= 0.66 Then
        ReDim Preserve tempbyte(i)
        Rnd -1
        Randomize Asc(Mid(Str(lngtotal), i, 1))
        tempbyte(i) = Int((90 - 65 + 1) * Rnd + 65)
        sRegNumber = sRegNumber & Chr(tempbyte(i))
    Else
        ReDim Preserve tempbyte(i)
        Rnd -1
        Randomize Asc(Mid(Str(lngtotal), i, 1))
        tempbyte(i) = Int((122 - 97 + 1) * Rnd + 97)
        sRegNumber = sRegNumber & Chr(tempbyte(i))
    End If
Next
txtRegNum = sRegNumber
End Function
