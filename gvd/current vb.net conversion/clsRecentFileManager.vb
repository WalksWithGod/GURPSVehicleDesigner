Option Strict Off
Option Explicit On
Friend Class clsRecentFileManager
	Dim m_oMenu() As System.Windows.Forms.ToolStripMenuItem
	
	Public WriteOnly Property Menu() As System.Windows.Forms.ToolStripMenuItem
		Set(ByVal Value As System.Windows.Forms.ToolStripMenuItem)
			ReDim m_oMenu(0)
			
			m_oMenu(0) = Value
			
			'test
			
		End Set
	End Property
	
	Function AddRecentFile(ByRef sFullPath As String) As Integer
		
		'//add this file to our number 1 spot.  If it is already at the number 1 spot
		'//move it down a notch and same with all below it except for 5th which gets bumped
		'//if the spot we are moving down to is the same as any of the ones above it, delete it
		Dim i As Object
		Dim j As Integer
		Dim sCurrentFile As String
		Dim sCurrentPath As String
		Dim sPrevFile As String
		Dim sPrevPath As String
		
		If Not m_oMenu(0) Is Nothing Then
			
			sCurrentFile = ExtractFileNameFromPath(sFullPath)
			sCurrentPath = ExtractPathFromFile(sFullPath)
			
			On Error Resume Next
			
			
			
			'//if its already on the recent file list, we wont add it again
			If Not AlreadyInList(sFullPath) Then
				For i = 5 To 2 Step -1
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					m_oMenu(i).Text = m_oMenu(i - 1).Text
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Debug.Print("AddRecentFile: " & "Caption " & i & " = " & m_oMenu(i).Text)
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Debug.Print("AddRecentFile: " & "Caption " & i - 1 & " = " & m_oMenu(i - 1).Text)
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					m_oMenu(i).Visible = True
					'mnuRecent(i).Caption = "&" & i & " " & Left(mnuRecent(i - 1).Caption, Len(mnuRecent(i - 1).Caption) - 3)
				Next 
				m_oMenu(1).Text = sFullPath
				'mnuRecent(1).Tag = sFilePath
				m_oMenu(1).Visible = True
				
				'//hide all the ones with "" for captions
				For i = 1 To 5
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If m_oMenu(i).Text = "" Then m_oMenu(i).Visible = False
				Next 
				
				'//make sure our seperator is visible
				If m_oMenu(1).Text <> "" Then m_oMenu(0).Visible = True
			End If
		End If
	End Function
	
	Sub DeleteRecentFile(ByRef sFilePath As String)
		Dim i As Integer
		
		For i = 1 To 5
			If m_oMenu(i).Text = sFilePath Then
				m_oMenu(i).Text = ""
				m_oMenu(i).Visible = False
			End If
		Next 
		
	End Sub
	Private Function AlreadyInList(ByRef sFilePath As Object) As Boolean
		Dim i As Integer
		
		For i = 5 To 1 Step -1
			'UPGRADE_WARNING: Couldn't resolve default property of object sFilePath. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If m_oMenu(i).Text = sFilePath Then
				AlreadyInList = True
				Exit Function
			End If
		Next 
		
		AlreadyInList = False
	End Function
	
	'UPGRADE_NOTE: Class_Terminate was upgraded to Class_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Terminate_Renamed()
		'UPGRADE_NOTE: Object m_oMenu() may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		m_oMenu(0) = Nothing
	End Sub
	Protected Overrides Sub Finalize()
		Class_Terminate_Renamed()
		MyBase.Finalize()
	End Sub
End Class