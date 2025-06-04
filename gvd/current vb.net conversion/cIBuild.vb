Option Strict Off
Option Explicit On
Public Interface _cIBuild
	Function getOption(ByVal lngIndex As Integer) As Integer
	Function setOption(ByVal lngIndex As Integer, ByVal lngSelection As Integer) As Boolean
	Function getUserInput(ByVal lngIndex As Integer) As Single
	Function setUserInput(ByVal lngIndex As Integer, ByVal sngValue As Single) As Boolean
End Interface
<System.Runtime.InteropServices.ProgId("cIBuild_NET.cIBuild")> Public Class cIBuild
	Implements _cIBuild
	' -------cIBuild related variables
	' NOTE: one benefit of having these in a seperate build interface is that unlike stats in cIcomponent
	' for weight/volume/cost etc, these build stats are not needed at run-time.  weight/volume/cost are needed
	' since these items can be bought/sold/refitted at run-time so we must retain this data at run-time.
	' NOTE1: I need to elaborate on the distinction between going to a vehicle "shop" and building / refitting a
	'        vehicle.  The "build" interface is only needed by the server and the client only  needs the end result
	'        so not having to load a build interface is good for the client AND for the server once its built too
	'        since active vehicles in memory wont take up as much space
	' NOTE2: Also, if we ever do switch to C++ or D, then we can share alot of code by just implementing
	' cIBuild once since I don't think we'd ever need to override it eh?
	' NOTE3: SHould cIBuild also be the interface that supports registering/unregistering of the global formula vars?
	
	
	' These tables are ONLY for formula calculation tables and NOT for modifiers.
	' Keep in mind that weapons for instance have a TON of formulas and we'll be treating every one of the
	' minor ones the same way as our major stats
	Private Structure GVD_DATA_TABLE
		Dim ID As Integer 'eg TABLE_COMPONENT_BUILD_DATA, TABLE_SCAN_RATING, TABLE_BOLTTHROWER_ST_DATA
		Dim ptrTable As Integer ' pointer to the TABLE UDT defined in modMatrix
	End Structure
	
	' the option tables hold the modifier settings for each selected option
	Private Structure GVD_OPTIONS
		Dim index As Integer ' index of the selected option in the drop down list.
		' from the index of the m_Options() and the index of
		' m_Options(x)index, we know the index of the m_Tables(?)
		' to use for the selection.  Does this need to be a UDT? why not
		' use an array of longs then?  Well, one m_Options() only uses
		' one data table regardless of the number of selections for that option
		' So dont we need GVD_OPTIONS so that in addition to index of the selected
		' we also know the number of selections there are?  Else how do we compute
		' the offset.  Further, having a numSelections will help us check that a user
		' doesnt try to set an index value that is invalid
		Dim selectionCount As Integer
		Dim ptrTable As Integer ' hrm...
	End Structure
	
	Private Structure GVD_USER_INPUT
		Dim sngValue As Single
		Dim sngURange As Single
		Dim sngLRange As Single
	End Structure
	
	Private Structure GVD_FORMULA
		Dim lngStatID As Integer ' id for cost,weight,surface,etc
		Dim lngFormulaID As Integer ' id of the formula to use
	End Structure
	
	Private m_Tables() As GVD_DATA_TABLE
	Private m_Options() As GVD_OPTIONS
	Private m_UserInput() As GVD_USER_INPUT
	Private m_Formulas() As GVD_FORMULA ' each stats (cost,weight,surface,etc) will need to be calc'd based on the ID of the formula to be used
	Private m_lngTableCount As Integer
	Private m_lngOptionCount As Integer 'the number of individual options and NOT the selections available for any specific option
	Private m_lngUserInputCount As Integer
	Private m_lngFormulaCount As Integer
	
	' todo: What about 'm_Bonus()' for advantages that a component gives?  Here or what?
	' No, m_Bonus() is needed at run-time so should be apart of cIComponent?
	Public Function getOption(ByVal lngIndex As Integer) As Integer Implements _cIBuild.getOption
	End Function
	Public Function setOption(ByVal lngIndex As Integer, ByVal lngSelection As Integer) As Boolean Implements _cIBuild.setOption
	End Function
	Public Function getUserInput(ByVal lngIndex As Integer) As Single Implements _cIBuild.getUserInput
	End Function
	Public Function setUserInput(ByVal lngIndex As Integer, ByVal sngValue As Single) As Boolean Implements _cIBuild.setUserInput
	End Function
	Private Function calcStats(ByRef oVisitor As cStats) As Boolean
	End Function
End Class