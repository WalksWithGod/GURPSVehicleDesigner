Attribute VB_Name = "modAutoReg"
Option Explicit


Sub Main()

    Dim sCommand As String
    Dim frm As frmAutoReg
    
    
    sCommand = Command()
    
    Set frm = New frmAutoReg
    frm.txtRegSoft = sCommand
    frm.Show
    

End Sub
