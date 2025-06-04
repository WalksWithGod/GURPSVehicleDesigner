VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIBuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
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
Private Type GVD_DATA_TABLE
    ID As Long   'eg TABLE_COMPONENT_BUILD_DATA, TABLE_SCAN_RATING, TABLE_BOLTTHROWER_ST_DATA
    ptrTable As Long ' pointer to the TABLE UDT defined in modMatrix
End Type

' the option tables hold the modifier settings for each selected option
Private Type GVD_OPTIONS
    index As Long ' index of the selected option in the drop down list.
                  ' from the index of the m_Options() and the index of
                  ' m_Options(x)index, we know the index of the m_Tables(?)
                  ' to use for the selection.  Does this need to be a UDT? why not
                  ' use an array of longs then?  Well, one m_Options() only uses
                  ' one data table regardless of the number of selections for that option
                  ' So dont we need GVD_OPTIONS so that in addition to index of the selected
                  ' we also know the number of selections there are?  Else how do we compute
                  ' the offset.  Further, having a numSelections will help us check that a user
                  ' doesnt try to set an index value that is invalid
    selectionCount As Long
    ptrTable As Long ' hrm...
End Type

Private Type GVD_USER_INPUT
    sngValue As Single
    sngURange As Single
    sngLRange As Single
End Type

Private Type GVD_FORMULA
    lngStatID As Long ' id for cost,weight,surface,etc
    lngFormulaID As Long ' id of the formula to use
End Type

Private m_Tables() As GVD_DATA_TABLE
Private m_Options() As GVD_OPTIONS
Private m_UserInput() As GVD_USER_INPUT
Private m_Formulas() As GVD_FORMULA ' each stats (cost,weight,surface,etc) will need to be calc'd based on the ID of the formula to be used
Private m_lngTableCount As Long
Private m_lngOptionCount As Long 'the number of individual options and NOT the selections available for any specific option
Private m_lngUserInputCount As Long
Private m_lngFormulaCount As Long

' todo: What about 'm_Bonus()' for advantages that a component gives?  Here or what?
' No, m_Bonus() is needed at run-time so should be apart of cIComponent?
Public Function getOption(ByVal lngIndex As Long) As Long
End Function
Public Function setOption(ByVal lngIndex As Long, ByVal lngSelection As Long) As Boolean
End Function
Public Function getUserInput(ByVal lngIndex As Long) As Single
End Function
Public Function setUserInput(ByVal lngIndex As Long, ByVal sngValue As Single) As Boolean
End Function
Private Function calcStats(ByRef oVisitor As cStats) As Boolean
End Function
