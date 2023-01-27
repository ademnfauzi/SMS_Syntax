Dim dt_TCID, dt_ScenarioName, dt_TestCase, dt_ExpectedResult

Call spLoadLibrary()
Call spInitiateData("Excel_Report.xlsx", "SCIC0002 - Create DEP account on same CIF.xlsx", "sheetname")
Call spGetDatatable()
Call fnRunningIterator()
Call spReportInitiate()
Call spAddScenario(dt_TCID, dt_ScenarioName, dt_TestCase, dt_ExpectedResult, "")


Call DA_Login_ICONS()
Call GoTo_ScreenNumber()
Call Searching_Nasabah_62000()
Call Filling_Mandatory_Field_2000()
Call Filling_Mandatory_Field_500()
Call DA_Logout_ICONS()


Call spReportSave() @@ hightlight id_;_Browser("ICONS 64 Alternity").Page("ICONS 64 Alternity").Frame("MainScreen").WebList("notipros")_;_script infofile_;_ZIP::ssf25.xml_;_
 @@ hightlight id_;_Browser("ICONS 64 Alternity").Page("ICONS 64 Alternity").Frame("MainScreen").WebList("OpenAcctRes")_;_script infofile_;_ZIP::ssf5.xml_;_

REM ========== SUB LOAD LIBRARY
Sub spLoadLibrary()
	Dim objSysInfo, Path_Env, LibFunction
	Set objSysInfo 		= Createobject("Wscript.Network")	
	Path_Env = Environment.Value("Path_Folder")
	LibFunction = Path_Env & "Libraries\"
	LibRepo = Path_Env & "Repositories\"
	LoadFunctionLibrary (LibFunction & "Lib_Report.vbs")
	LoadFunctionLibrary (LibFunction & "Lib_GlobalFunction.qfl")
	
	LoadFunctionLibrary (LibFunction & "Lib_BitviseLog.qfl")
	LoadFunctionLibrary (LibFunction & "Lib_BitviseTerminal.qfl")
	LoadFunctionLibrary (LibFunction & "Lib_Screen_62000.qfl")
	LoadFunctionLibrary (LibFunction & "Lib_ICONS.qfl")
	
	Call RepositoriesCollection.Add(LibRepo & "RP_Bitvice.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_ICONS_Main.tsr")
	Call RepositoriesCollection.Add(LibRepo & "RP_Screen_62000.tsr")
	
End Sub

Sub spGetDatatable()
	REM ---------- Report Data
	dt_TCID					= DataTable.Value("TC_ID", dtLocalSheet)
	dt_ScenarioName			= DataTable.Value("SCENARIO_NAME", dtLocalSheet)
	dt_TestCase				= DataTable.Value("TEST_CASE", dtLocalSheet)
	dt_ExpectedResult		= DataTable.Value("EXPECTED_RESULT", dtLocalSheet)
	
End Sub

