VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHTMLViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' HTML.BAS is a library for generating HTML
' files from Visual Basic.  This version only
' generates a small subset of HTML, but it should
' be useful for small applications, to help you
' learn HTML, or as a starting point on a more
' complete library.
'
' I welcome any questions or comments.  Send them
' to Andy Dean (andyd@mindspring.com)
'
' There are several classes of procedures:
' Function htmlEmbed(stag as string, ByRef s as string) as String
'   - Embeds a string in the appropriate
'     surrounding tags, and returns the
'     resulting string.
' ex: Bold(ByRef s as string)
'     The statement Bold("foo") generates
'         <bold>foo</bold>

' StartX, EndX
' These procedures define larger html structures
' or containers which include other data.
' These could be implemented the same as the
' htmlEmbed() family, but we are not writing LISP
' here!
'  Sub StartHTML(sfile as string)
'     Opens the file for output, and
'     emits the <HTML> tag.
'  Sub EndHTML()
'     Emits the closing </HTML> tag
'     and closes the file.
'
'
'  StartList( sType as string)
'      Emits the <OL> or <UL> as appropriate.
'
'  EndList()
'      Emits the closing </OL> or </UL> as appropriate.
'      (Use a stack to use the appropriate tag.  Don;t
'      need to specify the type of list.
'  StartTable
'  EndTable
'
'  Sub Emit()
'      Actually writes text to the target file.
'
'  Should perform all sort of testing, like insuring
'  that heading levels are valid, things that
'  end also have a beginning, or that all
'  tags are actually closed when a section ends.
'
'  Should probably implement this as a class,
'  so that no source code shows.
'
'  Will need to include a good stack utility to
'  manage the tests for correct nesting.
'
' Copyright (C)  1996-1997 Andy Dean
'

Private m_hFile As Integer

Function htmlBold(ByRef s As String)
   htmlBold = htmlEmbed("bold", s)
End Function

Function htmlCenter(ByRef s As String)
   htmlCenter = htmlEmbed("center", s)
End Function
Function htmlItalics(ByRef s As String)
   htmlItalics = htmlEmbed("i", s)
End Function


Sub htmlEmit(ByRef s As String)
  Debug.Print s
  'Emit s
  Print #m_hFile, s
End Sub

Sub htmlEmitStart(sTag As String)
  htmlEmit "<" & sTag & ">"
End Sub

Sub htmlEmitEnd(sTag As String)
  htmlEmit "</" & sTag & ">"
End Sub

Sub htmlEmitComment(ByRef s As String)
   htmlEmit "<! " & s & ">"
End Sub

Sub htmlEmitHREF(protocol As String, url As String, label As String)

  Dim QUOTE As String
  QUOTE = Chr$(34)
  
  Dim colon As String
   
  Select Case LCase$(protocol)
     Case "http"
        colon = "://"
     Case "mailto"
        colon = ":"
     Case "news"
        colon = ":"
     Case Else
        colon = ":"
  End Select
  htmlEmit "<A href=" & QUOTE & protocol & colon & url & QUOTE & ">" & label & "</a>"
End Sub

Sub htmlEmitImage(url As String, label As String)
  htmlEmit "Not implemented"
End Sub

Sub htmlEndHead()
  htmlEmitEnd "HEAD"
End Sub

Sub htmlComment(ByRef s As String)
  htmlEmit "<! " & s & ">"
End Sub
Sub htmlHR()
   htmlEmit "<HR>"
End Sub

Sub htmlBR()
   htmlEmit "<BR>"
End Sub

Sub htmlStartBody()
   htmlEmitStart "BODY"
End Sub

Sub htmlEndBody()
   htmlEmitEnd "BODY"
End Sub

Sub htmlEmitEmbedded(sTag As String, sText As String)
   htmlEmit htmlEmbed(sTag, sText)
End Sub

Sub htmlEmitHeading(iLevel As Integer, sText As String)
   Dim sTag As String
   sTag = "H" & iLevel
   htmlEmitEmbedded sTag, sText
End Sub

Sub htmlImage(sURL As String, sAlt As String)
   Dim QUOTE As String
   QUOTE = Chr$(34)
   htmlEmit "<IMG src=" & QUOTE & sURL & QUOTE & " alt=" & QUOTE & sAlt & QUOTE & ">"
End Sub

Sub htmlStartHead(sTitle As String)
   htmlEmitStart "HEAD"
   htmlEmitEmbedded "TITLE", sTitle      ' sTitle is required.
End Sub

Public Sub htmlStartHTML(sFileName)
   
  ' Open the file
 
  m_hFile = FreeFile
  Open sFileName For Output As #m_hFile
      
  ' Emit the <HTML> tag.
  htmlEmitStart "HTML"
  htmlEmitComment "Generated using HTML.BAS, by Andy Dean (andyd@mindspring.com)"

End Sub

Sub htmlEmitLineBreak()
   htmlEmit "<BR>"
End Sub

Sub htmlParagraph()
   htmlEmit "<P>"
End Sub

Public Sub htmlEndHTML()

    ' This should really check to make sure that

    ' Emit the </HTML> tag.
    htmlEmitEnd "HTML"
    
    ' Close the file
    Close #m_hFile
    
End Sub

Function htmlEmbed(attrib As String, ByRef s As String)
   htmlEmbed = "<" & attrib & ">" & s & "</" & attrib & ">"
End Function

Sub htmlListStart(sListType As String)
   htmlEmit "<" & sListType & "L>"
End Sub

Sub htmlListItem(sText As String)
   htmlEmit "<LI> " & sText
End Sub

Sub htmlListEnd(sListType As String)
   htmlEmit "</" & sListType & "L>"
End Sub

Sub htmlTableStart()
  htmlEmit "<TABLE border=1>"
End Sub

Sub htmlTableEnd()
  htmlEmitEnd "TABLE"
End Sub

Sub htmlTableRowStart()
  htmlEmitStart "TR"
End Sub

Sub htmlTableRowEnd()
  htmlEmitEnd "TR"
End Sub

Sub htmlTableData(sCell As String)
   htmlEmitEmbedded "TD", sCell
End Sub

Sub htmlTableHeader(sCell As String)
   htmlEmitEmbedded "TH", sCell
End Sub


' Display a
Public Sub htmlTabulateWeaponGun(ByVal lngNumFields As Long, ByRef sDelimitedString() As String, ByRef sDelimiter As String)

  Dim i  As Long
  Dim lngNumRows As Long
  Dim sFields() As String
  Dim j As Long
  lngNumRows = UBound(sDelimitedString)
  
   
   htmlTableStart
   htmlTableRowStart
   
   sFields() = Split(sDelimitedString(i))
   For i = 0 To lngNumFields - 1
      htmlTableHeader sFields(i)
   Next
   
   htmlTableRowEnd
   
   For i = 1 To lngNumRows - 1
      htmlTableRowStart
      sFields = Split(sDelimitedString(i))
      For j = 0 To lngNumFields - 1
         htmlTableData sFields(j)
      Next
      htmlTableRowEnd
   Next
   
   htmlTableEnd
   
End Sub

