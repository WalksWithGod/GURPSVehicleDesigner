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
vbwProfiler.vbwProcIn 379
vbwProfiler.vbwExecuteLine 9438
   htmlBold = htmlEmbed("bold", s)
vbwProfiler.vbwProcOut 379
vbwProfiler.vbwExecuteLine 9439
End Function

Function htmlCenter(ByRef s As String)
vbwProfiler.vbwProcIn 380
vbwProfiler.vbwExecuteLine 9440
   htmlCenter = htmlEmbed("center", s)
vbwProfiler.vbwProcOut 380
vbwProfiler.vbwExecuteLine 9441
End Function
Function htmlItalics(ByRef s As String)
vbwProfiler.vbwProcIn 381
vbwProfiler.vbwExecuteLine 9442
   htmlItalics = htmlEmbed("i", s)
vbwProfiler.vbwProcOut 381
vbwProfiler.vbwExecuteLine 9443
End Function


Sub htmlEmit(ByRef s As String)
vbwProfiler.vbwProcIn 382
vbwProfiler.vbwExecuteLine 9444
  Debug.Print s
  'Emit s
vbwProfiler.vbwExecuteLine 9445
  Print #m_hFile, s
vbwProfiler.vbwProcOut 382
vbwProfiler.vbwExecuteLine 9446
End Sub

Sub htmlEmitStart(sTag As String)
vbwProfiler.vbwProcIn 383
vbwProfiler.vbwExecuteLine 9447
  htmlEmit "<" & sTag & ">"
vbwProfiler.vbwProcOut 383
vbwProfiler.vbwExecuteLine 9448
End Sub

Sub htmlEmitEnd(sTag As String)
vbwProfiler.vbwProcIn 384
vbwProfiler.vbwExecuteLine 9449
  htmlEmit "</" & sTag & ">"
vbwProfiler.vbwProcOut 384
vbwProfiler.vbwExecuteLine 9450
End Sub

Sub htmlEmitComment(ByRef s As String)
vbwProfiler.vbwProcIn 385
vbwProfiler.vbwExecuteLine 9451
   htmlEmit "<! " & s & ">"
vbwProfiler.vbwProcOut 385
vbwProfiler.vbwExecuteLine 9452
End Sub

Sub htmlEmitHREF(protocol As String, url As String, label As String)
vbwProfiler.vbwProcIn 386

  Dim QUOTE As String
vbwProfiler.vbwExecuteLine 9453
  QUOTE = Chr$(34)

  Dim colon As String

vbwProfiler.vbwExecuteLine 9454
  Select Case LCase$(protocol)
'vbwLine 9455:     Case "http"
     Case IIf(vbwProfiler.vbwExecuteLine(9455), VBWPROFILER_EMPTY, _
        "http")
vbwProfiler.vbwExecuteLine 9456
        colon = "://"
'vbwLine 9457:     Case "mailto"
     Case IIf(vbwProfiler.vbwExecuteLine(9457), VBWPROFILER_EMPTY, _
        "mailto")
vbwProfiler.vbwExecuteLine 9458
        colon = ":"
'vbwLine 9459:     Case "news"
     Case IIf(vbwProfiler.vbwExecuteLine(9459), VBWPROFILER_EMPTY, _
        "news")
vbwProfiler.vbwExecuteLine 9460
        colon = ":"
     Case Else
vbwProfiler.vbwExecuteLine 9461 'B
vbwProfiler.vbwExecuteLine 9462
        colon = ":"
  End Select
vbwProfiler.vbwExecuteLine 9463 'B
vbwProfiler.vbwExecuteLine 9464
  htmlEmit "<A href=" & QUOTE & protocol & colon & url & QUOTE & ">" & label & "</a>"
vbwProfiler.vbwProcOut 386
vbwProfiler.vbwExecuteLine 9465
End Sub

Sub htmlEmitImage(url As String, label As String)
vbwProfiler.vbwProcIn 387
vbwProfiler.vbwExecuteLine 9466
  htmlEmit "Not implemented"
vbwProfiler.vbwProcOut 387
vbwProfiler.vbwExecuteLine 9467
End Sub

Sub htmlEndHead()
vbwProfiler.vbwProcIn 388
vbwProfiler.vbwExecuteLine 9468
  htmlEmitEnd "HEAD"
vbwProfiler.vbwProcOut 388
vbwProfiler.vbwExecuteLine 9469
End Sub

Sub htmlComment(ByRef s As String)
vbwProfiler.vbwProcIn 389
vbwProfiler.vbwExecuteLine 9470
  htmlEmit "<! " & s & ">"
vbwProfiler.vbwProcOut 389
vbwProfiler.vbwExecuteLine 9471
End Sub
Sub htmlHR()
vbwProfiler.vbwProcIn 390
vbwProfiler.vbwExecuteLine 9472
   htmlEmit "<HR>"
vbwProfiler.vbwProcOut 390
vbwProfiler.vbwExecuteLine 9473
End Sub

Sub htmlBR()
vbwProfiler.vbwProcIn 391
vbwProfiler.vbwExecuteLine 9474
   htmlEmit "<BR>"
vbwProfiler.vbwProcOut 391
vbwProfiler.vbwExecuteLine 9475
End Sub

Sub htmlStartBody()
vbwProfiler.vbwProcIn 392
vbwProfiler.vbwExecuteLine 9476
   htmlEmitStart "BODY"
vbwProfiler.vbwProcOut 392
vbwProfiler.vbwExecuteLine 9477
End Sub

Sub htmlEndBody()
vbwProfiler.vbwProcIn 393
vbwProfiler.vbwExecuteLine 9478
   htmlEmitEnd "BODY"
vbwProfiler.vbwProcOut 393
vbwProfiler.vbwExecuteLine 9479
End Sub

Sub htmlEmitEmbedded(sTag As String, sText As String)
vbwProfiler.vbwProcIn 394
vbwProfiler.vbwExecuteLine 9480
   htmlEmit htmlEmbed(sTag, sText)
vbwProfiler.vbwProcOut 394
vbwProfiler.vbwExecuteLine 9481
End Sub

Sub htmlEmitHeading(iLevel As Integer, sText As String)
vbwProfiler.vbwProcIn 395
   Dim sTag As String
vbwProfiler.vbwExecuteLine 9482
   sTag = "H" & iLevel
vbwProfiler.vbwExecuteLine 9483
   htmlEmitEmbedded sTag, sText
vbwProfiler.vbwProcOut 395
vbwProfiler.vbwExecuteLine 9484
End Sub

Sub htmlImage(sURL As String, sAlt As String)
vbwProfiler.vbwProcIn 396
   Dim QUOTE As String
vbwProfiler.vbwExecuteLine 9485
   QUOTE = Chr$(34)
vbwProfiler.vbwExecuteLine 9486
   htmlEmit "<IMG src=" & QUOTE & sURL & QUOTE & " alt=" & QUOTE & sAlt & QUOTE & ">"
vbwProfiler.vbwProcOut 396
vbwProfiler.vbwExecuteLine 9487
End Sub

Sub htmlStartHead(sTitle As String)
vbwProfiler.vbwProcIn 397
vbwProfiler.vbwExecuteLine 9488
   htmlEmitStart "HEAD"
vbwProfiler.vbwExecuteLine 9489
   htmlEmitEmbedded "TITLE", sTitle      ' sTitle is required.
vbwProfiler.vbwProcOut 397
vbwProfiler.vbwExecuteLine 9490
End Sub

Public Sub htmlStartHTML(sFileName)
vbwProfiler.vbwProcIn 398

  ' Open the file

vbwProfiler.vbwExecuteLine 9491
  m_hFile = FreeFile
vbwProfiler.vbwExecuteLine 9492
  Open sFileName For Output As #m_hFile

  ' Emit the <HTML> tag.
vbwProfiler.vbwExecuteLine 9493
  htmlEmitStart "HTML"
vbwProfiler.vbwExecuteLine 9494
  htmlEmitComment "Generated using HTML.BAS, by Andy Dean (andyd@mindspring.com)"

vbwProfiler.vbwProcOut 398
vbwProfiler.vbwExecuteLine 9495
End Sub

Sub htmlEmitLineBreak()
vbwProfiler.vbwProcIn 399
vbwProfiler.vbwExecuteLine 9496
   htmlEmit "<BR>"
vbwProfiler.vbwProcOut 399
vbwProfiler.vbwExecuteLine 9497
End Sub

Sub htmlParagraph()
vbwProfiler.vbwProcIn 400
vbwProfiler.vbwExecuteLine 9498
   htmlEmit "<P>"
vbwProfiler.vbwProcOut 400
vbwProfiler.vbwExecuteLine 9499
End Sub

Public Sub htmlEndHTML()
vbwProfiler.vbwProcIn 401

    ' This should really check to make sure that

    ' Emit the </HTML> tag.
vbwProfiler.vbwExecuteLine 9500
    htmlEmitEnd "HTML"

    ' Close the file
vbwProfiler.vbwExecuteLine 9501
    Close #m_hFile

vbwProfiler.vbwProcOut 401
vbwProfiler.vbwExecuteLine 9502
End Sub

Function htmlEmbed(attrib As String, ByRef s As String)
vbwProfiler.vbwProcIn 402
vbwProfiler.vbwExecuteLine 9503
   htmlEmbed = "<" & attrib & ">" & s & "</" & attrib & ">"
vbwProfiler.vbwProcOut 402
vbwProfiler.vbwExecuteLine 9504
End Function

Sub htmlListStart(sListType As String)
vbwProfiler.vbwProcIn 403
vbwProfiler.vbwExecuteLine 9505
   htmlEmit "<" & sListType & "L>"
vbwProfiler.vbwProcOut 403
vbwProfiler.vbwExecuteLine 9506
End Sub

Sub htmlListItem(sText As String)
vbwProfiler.vbwProcIn 404
vbwProfiler.vbwExecuteLine 9507
   htmlEmit "<LI> " & sText
vbwProfiler.vbwProcOut 404
vbwProfiler.vbwExecuteLine 9508
End Sub

Sub htmlListEnd(sListType As String)
vbwProfiler.vbwProcIn 405
vbwProfiler.vbwExecuteLine 9509
   htmlEmit "</" & sListType & "L>"
vbwProfiler.vbwProcOut 405
vbwProfiler.vbwExecuteLine 9510
End Sub

Sub htmlTableStart()
vbwProfiler.vbwProcIn 406
vbwProfiler.vbwExecuteLine 9511
  htmlEmit "<TABLE border=1>"
vbwProfiler.vbwProcOut 406
vbwProfiler.vbwExecuteLine 9512
End Sub

Sub htmlTableEnd()
vbwProfiler.vbwProcIn 407
vbwProfiler.vbwExecuteLine 9513
  htmlEmitEnd "TABLE"
vbwProfiler.vbwProcOut 407
vbwProfiler.vbwExecuteLine 9514
End Sub

Sub htmlTableRowStart()
vbwProfiler.vbwProcIn 408
vbwProfiler.vbwExecuteLine 9515
  htmlEmitStart "TR"
vbwProfiler.vbwProcOut 408
vbwProfiler.vbwExecuteLine 9516
End Sub

Sub htmlTableRowEnd()
vbwProfiler.vbwProcIn 409
vbwProfiler.vbwExecuteLine 9517
  htmlEmitEnd "TR"
vbwProfiler.vbwProcOut 409
vbwProfiler.vbwExecuteLine 9518
End Sub

Sub htmlTableData(sCell As String)
vbwProfiler.vbwProcIn 410
vbwProfiler.vbwExecuteLine 9519
   htmlEmitEmbedded "TD", sCell
vbwProfiler.vbwProcOut 410
vbwProfiler.vbwExecuteLine 9520
End Sub

Sub htmlTableHeader(sCell As String)
vbwProfiler.vbwProcIn 411
vbwProfiler.vbwExecuteLine 9521
   htmlEmitEmbedded "TH", sCell
vbwProfiler.vbwProcOut 411
vbwProfiler.vbwExecuteLine 9522
End Sub


' Display a
Public Sub htmlTabulateWeaponGun(ByVal lngNumFields As Long, ByRef sDelimitedString() As String, ByRef sDelimiter As String)
vbwProfiler.vbwProcIn 412

  Dim i  As Long
  Dim lngNumRows As Long
  Dim sFields() As String
  Dim j As Long
vbwProfiler.vbwExecuteLine 9523
  lngNumRows = UBound(sDelimitedString)


vbwProfiler.vbwExecuteLine 9524
   htmlTableStart
vbwProfiler.vbwExecuteLine 9525
   htmlTableRowStart

vbwProfiler.vbwExecuteLine 9526
   sFields() = Split(sDelimitedString(i))
vbwProfiler.vbwExecuteLine 9527
   For i = 0 To lngNumFields - 1
vbwProfiler.vbwExecuteLine 9528
      htmlTableHeader sFields(i)
vbwProfiler.vbwExecuteLine 9529
   Next

vbwProfiler.vbwExecuteLine 9530
   htmlTableRowEnd

vbwProfiler.vbwExecuteLine 9531
   For i = 1 To lngNumRows - 1
vbwProfiler.vbwExecuteLine 9532
      htmlTableRowStart
vbwProfiler.vbwExecuteLine 9533
      sFields = Split(sDelimitedString(i))
vbwProfiler.vbwExecuteLine 9534
      For j = 0 To lngNumFields - 1
vbwProfiler.vbwExecuteLine 9535
         htmlTableData sFields(j)
vbwProfiler.vbwExecuteLine 9536
      Next
vbwProfiler.vbwExecuteLine 9537
      htmlTableRowEnd
vbwProfiler.vbwExecuteLine 9538
   Next

vbwProfiler.vbwExecuteLine 9539
   htmlTableEnd

vbwProfiler.vbwProcOut 412
vbwProfiler.vbwExecuteLine 9540
End Sub


