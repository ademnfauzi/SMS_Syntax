Public GlobalRowIndex
Dim testCaseName
testCaseName = Environment("TestName")

Function fnGetDataTable(byval filename)
	Dim fullpath_excel, fso_excel
	Dim Path_Env, Path_Ver

	Set fullpath_excel 	= Createobject("Wscript.Network")
	Set fso_excel 		= CreateObject("Scripting.FileSystemObject")
	
	Path_Env			= Environment.Value("Path_Folder")
	fullpath_excel 		= Path_Env & "Excel\" & filename
	fnGetDataTable 		= fullpath_excel
	
'	If fso.FileExists(fullpath) Then
'		Reporter.ReportEvent micWarning, "Source Data is not exist", fullpath
'		Call ExitTest()
'	End If
End Function

Sub spInitiateData(byval globalData, byval localData, ByVal sheetName)
REM ------ INITIATE EXCEL FILE
	On Error Resume Next
	Dim tempSheet
	Set tempSheet = DataTable.GetSheet("TEMPORARY")
	If tempSheet Is Nothing Then
	DataTable.AddSheet("TEMPORARY")
'	DataTable.AddSheet("API_CONFIG")
	DataTable.AddSheet("LOG")
'	DataTable.AddSheet("API_DB")
	DataTable.AddSheet("SMS")
	Dim dtGlobal, dtLocal, dtDBConfig, dtConfig
		
	dtGlobal 	= fnGetDataTable(globalData)
	Call DataTable.ImportSheet(dtGlobal, "Global", dtGlobalSheet)
	
	dtLocal = fnGetDataTable(localData)
	Call DataTable.ImportSheet(dtLocal, sheetName, dtLocalSheet)
		
	dtConfig = fnGetDataTable("Excel_Config.xlsx")
'	Call DataTable.ImportSheet(dtConfig, "API_CONFIG", "API_CONFIG")
	Call DataTable.ImportSheet(dtConfig, "LOG", "LOG")
'	Call DataTable.ImportSheet(dtConfig, "API_DB", "API_DB")
	Call DataTable.ImportSheet(dtConfig, "SMS", "SMS")
'	
	End If
	On Error GoTo 0
End Sub

Function fnRunningIterator()
	If CInt(Environment("ActionIteration")) = CInt(DataTable.LocalSheet.GetRowCount()) Then '== 4
		If Trim(DataTable.Value("RUN", dtLocalSheet)) = "" Then 'Statenya ga run
			spReportForceSave() 'Save
			ExitActionIteration()
			Exit Function	
		End If
	End If
	
	If Trim(DataTable.Value("RUN", dtLocalSheet)) = "" Then 'Statenya ga run
		ExitActionIteration()
		Exit Function	
	End If
End Function

Function spReportInitiate()
	Dim author, tester, shortDescHeader, shortDescBody
	Dim projectType, projectName, projectCode
	Dim coverTitle, coverSubTitle
	Dim Tester1, Tester2, TestManager, TestingGroupHead, DevelopmentManager, RequirementManager, ProjectManager
	
	author					= DataTable.Value ("AUTHOR", dtGlobalSheet)
	shortDescHeader			= DataTable.Value ("HEADER_DESCRIPTION", dtGlobalSheet)
	shortDescBody			= DataTable.Value ("HEADER_BODY", dtGlobalSheet)
	
	projectType				= DataTable.Value ("PROJECT_TYPE", dtGlobalSheet)
	projectName				= DataTable.Value ("PROJECT_NAME", dtGlobalSheet)
	projectCode				= DataTable.Value ("PROJECT_CODE", dtGlobalSheet)
		
	coverTitle				= DataTable.Value ("COVER_TITLE", dtGlobalSheet)
	coverSubTitle			= DataTable.Value ("COVER_SUBTITLE", dtGlobalSheet)
	
	Tester1					= DataTable.Value ("TESTER1", dtGlobalSheet)
	Tester2					= DataTable.Value ("TESTER2", dtGlobalSheet)
	TestManager				= DataTable.Value ("TEST_MANAGER", dtGlobalSheet)
	TestingGroupHead		= DataTable.Value ("TESTING_GROUPHEAD", dtGlobalSheet)
	DevelopmentManager	= DataTable.Value ("DEVELOPMENT_MANAGER", dtGlobalSheet)
	RequirementManager		= DataTable.Value ("REQUIREMENT_MANAGER", dtGlobalSheet)
	ProjectManager			= DataTable.Value ("PROJECT_MANAGER", dtGlobalSheet)
	
	REM ------------ Initiated Report Library
	Call spInitiateReport("Prepared By " & author, author, shortDescHeader, shortDescBody, "Reporting")
	Call spInitiateReportProject(projectType, projectName, projectCode)
	Call spInitiateReportCover(coverTitle, coverSubTitle)
	Call spInitiateReportSigner(Tester1, Tester2, TestManager, TestingGroupHead, DevelopmentManager, RequirementManager, ProjectManager)
	Call spInitiateReportAttributes()
'	Call spInitiateReportBusinessRequirements()
' 	Call spInitiateReportSystemImpacted()
'	Call spInitiateReportSystemChanges()
End Function



















REM =================================
REM Declare Project Type constant
REM =================================
Const M_CR = "CR" 'change request
Const M_IR = "IR" 
Const M_MR = "MR"

REM =================================
REM Declare another constant
REM =================================
Const M_NewLineCr = "\r"
Const M_NewLineLf = "\n"
Const M_NewLineCrLf = "\r\n"

REM =================================
REM Create a enumerable blueprint
REM =================================
Class StepStatusEntity
    Public  Passed, Warning, Failed, Done
    Private Sub Class_Initialize
        Passed = 0
        Warning = 1
        Failed = 2
		Done = 3
    End Sub
End Class

Class StepFormatEntity
    Public  FormatText, FormatJson, FormatXml
    Private Sub Class_Initialize
        FormatText = 0
        FormatJson = 1
        FormatXml = 2
    End Sub
End Class

Class CompatibilityEntity
    Public  Desktop, Mobile
    Private Sub Class_Initialize
        Desktop = 0
        Mobile = 1
    End Sub
End Class

REM =================================
REM Declare variables
REM =================================
Dim ReportStatus
Dim StepFormat
Dim CompatibilityMode

Dim PDI_ShortDescriptionHeader
Dim PDI_ShortDescriptionBody
Dim PDI_DocumentAuthor
Dim PDI_TesterName
Dim PDI_ProjectName
Dim PDI_ProjectCode
Dim PDI_ProjectType
Dim PDI_CoverTitle
Dim PDI_CoverSubTitle



REM =================================
REM Initiate all variables
REM =================================
Dim react
Public initiated
Public attributesHaveBeenSet

initiated = False
attributesHaveBeenSet = False

Set react = Nothing
Set ReportStatus = New StepStatusEntity
Set StepFormat = New StepFormatEntity
Set CompatibilityMode = New CompatibilityEntity

PDI_ShortDescriptionHeader = ""
PDI_ShortDescriptionBody = ""
PDI_CoverTitle = ""
PDI_CoverSubTitle = ""
PDI_DocumentAuthor = ""
PDI_TesterName = ""
PDI_ProjectType = ""
PDI_ProjectName = ""
PDI_ProjectCode = ""

REM =========================================
REM Declare to create new and set path folders
REM =========================================

Function LPad (str,pad,length)
	LPad = String(length-Len(str),pad) & str
End Function

Function fnCreateFolder(byval folderProject)
	Dim LibPathKeagenan, LibRepo, objSysInfo
	Dim Path_Env
	Dim folderMyDoc
	Dim folderPath
	Dim folderActionPath
	
	Set objSysInfo 		= Createobject("Wscript.Network")		
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	Path_Env	= Environment.Value("Path_Folder")
	folderMyDoc = Path_Env
	folderPath = folderMyDoc & "\" & folderProject
	
	RunReportDate = Year(Now) & LPad(Month(Date), "0", 2) & LPad(Day(Date), "0", 2)
	folderActionPath = folderPath & "\" & RunReportDate
	
	
	If Not objFSO.FolderExists(folderPath) Then
		Call objFSO.CreateFolder(folderPath)
	End If
	
	If Not objFSO.FolderExists(folderActionPath) Then
		Call objFSO.CreateFolder(folderActionPath)
	End If
	
	fnCreateFolder = folderActionPath
End Function

REM =========================================
REM Declare Aplic Report initiation functions
REM =========================================

REM *****************************************
REM Initiate Report Object
REM *****************************************
Sub spInitiateReport(ByVal author, ByVal tester, Byval shortDescHeader, ByVal shortDescBody, byval folderName)
	If Not initiated Then
		Set react = CreateObject("Aplic.Report")
		Call react.About()
		Call react.SetUseDocumentInfo(true)
		Call react.SetOutputDirectory(fnCreateFolder(folderName))
		Call react.SetImageSizePercentage(75)
		Call react.SetMoveWordDocumentAfterSave(true)
		Call react.SetScenarioPrefix("")
		Call react.SetDisplayDocumentDate(true)
		Call react.SetDisplayTestingDate(true)
		Call react.SetAutoHideWarningStatusIfEmptyInSummary(true)
		Call react.SetUseScenarioNumbering(false)
		Call react.SetDisplayCustomAttributes(true)
		Call react.SetDocumentAuthor(author)
		Call react.SetDocumentDate(fnGetDate())
		Call react.SetDocumentShortDescriptionHeader(shortDescHeader)
		Call react.SetDocumentShortDescriptionBody(shortDescBody)
		Call react.SetTester(tester)
		initiated = True
	End If
End Sub


REM *******************************************
REM Initiate Report Project Type, Name and Code
REM *******************************************
Sub spInitiateReportProject(ByVal projectType, ByVal projectName, ByVal projectCode)
	If initiated Then
		'Set react = CreateObject("Aplic.Report")
		Call react.AddProjectType(M_CR)
		Call react.AddProjectType(M_IR)
		Call react.AddProjectType(M_MR)
		Call react.SetProjectType(projectType)
		Call react.SetProjectName(projectName)
		Call react.SetProjectCode(projectCode)
	End If
End Sub

REM *************************************************
REM Initiate Report Cover
REM - To set cover image, refer to 
REM - "C:\ProgramData\Napalm\React\Napalm.Design.ini"
REM - Section [Cover], DocumentLogoPath
REM *************************************************
Sub spInitiateReportCover(ByVal coverTitle, ByVal coverSubTitle)
	Call react.SetCoverTitle(coverTitle)
	Call react.SetCoverSubTitle(coverSubTitle)
End Sub

REM *********************************************************
REM Initiate Report Signer
REM - Move this method/procedure to project specific library, 
REM - Since it will apply different for each project
REM 
REM Note: This information will be collected from customer
REM *********************************************************
Sub spInitiateReportSigner(byval Tester1, byval Tester2, byval TestManager, byval TestingGroupHead, byval DevelopmentManager, byval RequirementManager, byval ProjectManager)
	Call react.SetSigner(Tester1, "Tester/Developer")
	Call react.SetSigner(Tester2, "Tester/Developer")
	Call react.SetSigner(TestManager, "Test Manager")
	Call react.SetSigner(TestingGroupHead, "Testing Group Head")
	Call react.SetSigner(DevelopmentManager, "Development Manager")
	Call react.SetSigner(RequirementManager, "Requirement Manager/Business Analyst")
	Call react.SetSigner(ProjectManager, "Project Manager")
End Sub

REM *********************************************************
REM Initiate Report Attributes
REM - Move this method/procedure to project specific library, 
REM - Since it will apply different for each project
REM - Such as KeagenanGlobalLib.vbs (example lib name)
REM *********************************************************
Sub spInitiateReportAttributes()
	Dim attUFTVersion, attBrowser_Icons, attURL_Icons
	
	attUFTVersion				= DataTable.Value("UFT_VERSION", dtGlobalSheet)
'	attBrowser_Icons			= DataTable.Value("ICONS_BROWSER", "API_ICONS")
'	attURL_Icons				= DataTable.Value("ICONS_URL", "API_ICONS")
	
	Call react.SetCustomAttributes("UFT Version", attUFTVersion)
'	Call react.SetCustomAttributes("Browser Icons", attBrowser_Icons)
'	Call react.SetCustomAttributes("URL Icons", attURL_Icons)
	
	Call react.SetCustomAttributes("Report Library", "Lib_Report.vbs")
	Call react.SetCustomAttributes("Global Library","Lib_GlobalFunction.qfl")
	
'	Call react.SetCustomAttributes("API Global Library", "Lib_APIGlobal.qfl")
'	Call react.SetCustomAttributes("API Transaction Library", "Lib_API.vbs")
'	Call react.SetCustomAttributes("API Verification Library", "API_Verifications.qfl")
End Sub

REM *********************************************************
REM Initiate Report Impacted Systems
REM - Move this method/procedure to project specific library, 
REM - Since it will apply different for each project
REM 
REM Note: This information will be collected from customer
REM *********************************************************
Sub spInitiateReportSystemImpacted()
	Call react.AddSystemImpacted("")
	Call react.AddSystemImpacted("")
	Call react.AddSystemImpacted("")
End Sub


REM =================================
REM Declare Aplic Report functions
REM =================================
Sub spAddScenario(ByVal no, ByVal useCaseDesc, ByVal scenarioDesc, ByVal exitCriteria, ByVal preparations)
	If IsArray(preparations) Then
		Dim sPreparations
		Dim index
		
		For index = 0 to UBound(preparations)
			sPreparations = sPreparations & react.MultipleDelimiter & preparations(index)
		Next
		
		sPreparations = Mid(sPreparations, 2)
		Call react.AddTitle(no, useCaseDesc, scenarioDesc, exitCriteria, sPreparations)
	Else
		If VarType(preparations) = vbString Then
			Call react.AddTitle(no, useCaseDesc, scenarioDesc, exitCriteria, preparations)
		End If
	End If
End Sub

REM *********************************
REM Create Scenario
REM *********************************
'Sub spAddScenario(ByVal scenarioName, ByVal exitCriteria, ByVal preparations)
'	If IsArray(preparations) Then
'		Dim sPreparations
'		Dim index
'		
'		For index = 0 to UBound(preparations)
'			sPreparations = sPreparations & react.MultipleDelimiter & preparations(index)
'		Next
'		
'		sPreparations = Mid(sPreparations, 2)
'		Call react.AddTitle(scenarioName, exitCriteria, sPreparations)
'	Else
'		If VarType(preparations) = vbString Then
'			Call react.AddTitle(scenarioName, exitCriteria, preparations)
'		End If
'	End If
'End Sub

REM *********************************
REM Add Image Step to Scenario
REM *********************************
Sub spAddImage(ByVal stepName, ByVal imageDescription, ByVal imagePath, ByVal compatibilityMode, ByVal stepStatus)
	Call react.AddImage(stepName, imageDescription, imagePath, compatibilityMode, stepStatus)
End Sub

REM *********************************
REM Add Multiple Image Step to Scenario
REM imageList can be an array, use Array("image1Path", "image1Desc", "image2Path", "image2Desc")
REM imageList can be a string, use "image1Path|image1Desc|image2Path|image2Desc"
REM *********************************
Sub spAddImages(ByVal stepName, ByVal imageList, ByVal compatibilityMode, ByVal stepStatus)
	If IsArray(imageList) Then
		Dim sImageList
		Dim index
		
		For index = 0 to UBound(imageList) - 1
			sImageList = sImageList & react.MultipleDelimiter & imageList(index)
		Next
		
		sImageList = Mid(sImageList, 2)
		Call react.AddImages(stepName, sImageList, compatibilityMode, stepStatus)
	Else
		If VarType(imageList) = vbString Then
			Call react.AddImages(stepName, imageList, compatibilityMode, stepStatus)
		End If
	End If
End Sub

REM **********************************
REM Add Step to scenario without image
REM **********************************
Sub spAddStep(ByVal stepName, ByVal stepDescription, ByVal stepStatus)
	Call react.AddStep(stepName, stepDescription, stepStatus)
End Sub

REM **********************************
REM Add Formatted Step to scenario
REM Only support Text, JSON and XML
REM **********************************
Sub spAddFormattedStep(ByVal stepName, ByVal stepDescription, ByVal delimiter, ByVal stepFormat, ByVal stepStatus)
	Call react.AddFormattedStep(stepName, stepDescription, delimiter, stepFormat, stepStatus)
End Sub

REM ********************************************
REM Save Report and Generate
REM - Based on UFT Last Test Iteration
REM ********************************************
Sub spReportSave()
	If CInt(Environment("ActionIteration")) = CInt(DataTable.LocalSheet.GetRowCount()) Then
		If initiated Then
			react.Save Environment("ActionName"), true
			initiated = false
		End If
	End If
End Sub

REM ********************************************
REM Save Report and Generate
REM ********************************************
Sub spReportForceSave()
	If initiated Then
		react.Save Environment("ActionName"), true
		initated = false
	End If
End Sub
