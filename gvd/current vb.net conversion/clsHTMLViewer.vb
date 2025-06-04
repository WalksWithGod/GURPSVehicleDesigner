Option Strict Off
Option Explicit On
Friend Class clsHTMLViewer
	
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
	
	Private m_hFile As Short
	
	Function htmlBold(ByRef s As String) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object htmlEmbed(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		htmlBold = htmlEmbed("bold", s)
	End Function
	
	Function htmlCenter(ByRef s As String) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object htmlEmbed(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		htmlCenter = htmlEmbed("center", s)
	End Function
	Function htmlItalics(ByRef s As String) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object htmlEmbed(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		htmlItalics = htmlEmbed("i", s)
	End Function
	
	
	Sub htmlEmit(ByRef s As String)
		Debug.Print(s)
		'Emit s
		PrintLine(m_hFile, s)
	End Sub
	
	Sub htmlEmitStart(ByRef sTag As String)
		htmlEmit("<" & sTag & ">")
	End Sub
	
	Sub htmlEmitEnd(ByRef sTag As String)
		htmlEmit("</" & sTag & ">")
	End Sub
	
	Sub htmlEmitComment(ByRef s As String)
		htmlEmit("<! " & s & ">")
	End Sub
	
	Sub htmlEmitHREF(ByRef protocol As String, ByRef url As String, ByRef label As String)
		
		Dim QUOTE As String
		QUOTE = Chr(34)
		
		Dim colon As String
		
		Select Case LCase(protocol)
			Case "http"
				colon = "://"
			Case "mailto"
				colon = ":"
			Case "news"
				colon = ":"
			Case Else
				colon = ":"
		End Select
		htmlEmit("<A href=" & QUOTE & protocol & colon & url & QUOTE & ">" & label & "</a>")
	End Sub
	
	Sub htmlEmitImage(ByRef url As String, ByRef label As String)
		htmlEmit("Not implemented")
	End Sub
	
	Sub htmlEndHead()
		htmlEmitEnd("HEAD")
	End Sub
	
	Sub htmlComment(ByRef s As String)
		htmlEmit("<! " & s & ">")
	End Sub
	Sub htmlHR()
		htmlEmit("<HR>")
	End Sub
	
	Sub htmlBR()
		htmlEmit("<BR>")
	End Sub
	
	Sub htmlStartBody()
		htmlEmitStart("BODY")
	End Sub
	
	Sub htmlEndBody()
		htmlEmitEnd("BODY")
	End Sub
	
	Sub htmlEmitEmbedded(ByRef sTag As String, ByRef sText As String)
		'UPGRADE_WARNING: Couldn't resolve default property of object htmlEmbed(). Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		htmlEmit(htmlEmbed(sTag, sText))
	End Sub
	
	Sub htmlEmitHeading(ByRef iLevel As Short, ByRef sText As String)
		Dim sTag As String
		sTag = "H" & iLevel
		htmlEmitEmbedded(sTag, sText)
	End Sub
	
	Sub htmlImage(ByRef sURL As String, ByRef sAlt As String)
		Dim QUOTE As String
		QUOTE = Chr(34)
		htmlEmit("<IMG src=" & QUOTE & sURL & QUOTE & " alt=" & QUOTE & sAlt & QUOTE & ">")
	End Sub
	
	Sub htmlStartHead(ByRef sTitle As String)
		htmlEmitStart("HEAD")
		htmlEmitEmbedded("TITLE", sTitle) ' sTitle is required.
	End Sub
	
	Public Sub htmlStartHTML(ByRef sFileName As Object)
		
		' Open the file
		
		m_hFile = FreeFile
		'UPGRADE_WARNING: Couldn't resolve default property of object sFileName. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		FileOpen(m_hFile, sFileName, OpenMode.Output)
		
		' Emit the <HTML> tag.
		htmlEmitStart("HTML")
		htmlEmitComment("Generated using HTML.BAS, by Andy Dean (andyd@mindspring.com)")
		
	End Sub
	
	Sub htmlEmitLineBreak()
		htmlEmit("<BR>")
	End Sub
	
	Sub htmlParagraph()
		htmlEmit("<P>")
	End Sub
	
	Public Sub htmlEndHTML()
		
		' This should really check to make sure that
		
		' Emit the </HTML> tag.
		htmlEmitEnd("HTML")
		
		' Close the file
		FileClose(m_hFile)
		
	End Sub
	
	Function htmlEmbed(ByRef attrib As String, ByRef s As String) As Object
		'UPGRADE_WARNING: Couldn't resolve default property of object htmlEmbed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		htmlEmbed = "<" & attrib & ">" & s & "</" & attrib & ">"
	End Function
	
	Sub htmlListStart(ByRef sListType As String)
		htmlEmit("<" & sListType & "L>")
	End Sub
	
	Sub htmlListItem(ByRef sText As String)
		htmlEmit("<LI> " & sText)
	End Sub
	
	Sub htmlListEnd(ByRef sListType As String)
		htmlEmit("</" & sListType & "L>")
	End Sub
	
	Sub htmlTableStart()
		htmlEmit("<TABLE border=1>")
	End Sub
	
	Sub htmlTableEnd()
		htmlEmitEnd("TABLE")
	End Sub
	
	Sub htmlTableRowStart()
		htmlEmitStart("TR")
	End Sub
	
	Sub htmlTableRowEnd()
		htmlEmitEnd("TR")
	End Sub
	
	Sub htmlTableData(ByRef sCell As String)
		htmlEmitEmbedded("TD", sCell)
	End Sub
	
	Sub htmlTableHeader(ByRef sCell As String)
		htmlEmitEmbedded("TH", sCell)
	End Sub
	
	
	' Display a
	Public Sub htmlTabulateWeaponGun(ByVal lngNumFields As Integer, ByRef sDelimitedString() As String, ByRef sDelimiter As String)
		
		Dim i As Integer
		Dim lngNumRows As Integer
		Dim sFields() As String
		Dim j As Integer
		lngNumRows = UBound(sDelimitedString)
		
		
		htmlTableStart()
		htmlTableRowStart()
		
		sFields = Split(sDelimitedString(i))
		For i = 0 To lngNumFields - 1
			htmlTableHeader(sFields(i))
		Next 
		
		htmlTableRowEnd()
		
		For i = 1 To lngNumRows - 1
			htmlTableRowStart()
			sFields = Split(sDelimitedString(i))
			For j = 0 To lngNumFields - 1
				htmlTableData(sFields(j))
			Next 
			htmlTableRowEnd()
		Next 
		
		htmlTableEnd()
		
	End Sub
End Class