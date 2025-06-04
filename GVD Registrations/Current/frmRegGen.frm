VERSION 5.00
Begin VB.Form frmRegGen 
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
Attribute VB_Name = "frmRegGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' A pathetically bad registrion code generator for the GURPS Vehicle Designer program.  This was written
' ~ten years ago in about 30 minutes off the top of my head.  Keep in mind I'd only been programming
' for less than two years at that point.  Below is a stripped down version of the original code which
' contained allot of decoy routines and fake password variables that did nothing except hopefully make it more difficult for
' hackers to follow the algorithm when using a low level debugger.


' Command1_Click is an event responding to the button click on the form.
' it gets the Long value (32 bits.  This would be an Int32 or Integer in vb.net)
' and passes it along with the full user name (First Middle Last, ASCII letters. Spaces ok. No punctuation)
' to the GenerateRegCode function and then puts the results
Private Sub Command1_Click()
    Dim regID As Long
    regID = Val(txtRegID) 'this should be a unique "userID" of the reg'd user which I'd normally pull from a DB as the next record number

    txtRegNum.Text = GetRegCode(txtRegName.Text, regID)

End Sub


' Poor man's HASH function which in the end did the trick and as far as I know was never cracked, but then again
' not many crackers would be interested in cracking a GURPS utility :)
Public Function GetRegCode(ByVal userfullname As String, ByVal regID As Long) As String

    Dim byteRegName() As Byte
    Dim tempbyte() As Byte
    Dim i As Long
    Dim j As Single
    Dim sTName As String
    Dim lngtotal As Single ' not sure why i didnt make this a Long. I think it was to avoid a potential overflow
    Dim sRegNumber As String
    

    '1- convert the user's full name into a byte array
    '   and total the values of the odd valued ascii letters once, and the even ones twice.
    ReDim byteRegName(Len(userfullname))
    For i = 1 To Len(userfullname)
        byteRegName(i) = Asc(Mid(userfullname, i, 1))
        lngtotal = lngtotal + byteRegName(i)
        'at the same time total the ascii value for every even valued ascii code
        If byteRegName(i) Mod 2 = 0 Then
            lngtotal = lngtotal + byteRegName(i)
        End If
    Next
    
    '2 - the RegID is actually just a modifier to prevent two people who have the same
    '    name end up with the same ID.  This ID is unique and alone can be used
    '   to identify a user.  Multiply this to the running byte total
    lngtotal = lngtotal * regID
    
    '3- take the multiplied ascii value of the typename of each letter in our magic string and multiply that to it
    ' NOTE: In the actual GVD program, I retreive the magic string "clsBody" by calling Typename(targetObject)
    ' and did not simply have it in the clear as below.
    sTName = "clsBody"
    For i = 1 To Len(sTName)
        lngtotal = lngtotal * Asc(Mid(sTName, i, 1))
    Next
    '4- take a random seed to generate the seeded random number and multiply that
    Rnd -1
    Randomize 9921988 ' i dont remember why i originally decided to do this here, but it is now required to produce matching reg codes.
    lngtotal = lngtotal * Rnd()
    
    '5- return this as a byte array but we want to limit the range of the ascii
    '   values to be just string's so that our "regcode number" is made up only of letters
    '  So this means we have to limit our our ascii values to upper and lower case letters which means
    ' they must be between 48-57, 65-90 and 97-122
    ' So it looks like i decided to take the string representation of our lngTotal HASH value and itterate through
    ' each letter and get the ascii value and use it to seed the VB random number generator and then to use that to
    ' generate a known ascii upper or lower case letter.
    ReDim tempbyte(Len(Str(lngtotal)))
    For i = 1 To Len(Str(lngtotal))
        j = Rnd()
        Rnd -1
        Randomize Asc(Mid(Str(lngtotal), i, 1))
            
        If j <= 0.33 Then
            tempbyte(i) = Int((57 - 48 + 1) * Rnd + 48)
        ElseIf j <= 0.66 Then
            tempbyte(i) = Int((90 - 65 + 1) * Rnd + 65)
        Else
            tempbyte(i) = Int((122 - 97 + 1) * Rnd + 97)
        End If
        
        sRegNumber = sRegNumber & Chr(tempbyte(i))
    Next

    ' return
    GetRegCode = sRegNumber
End Function


