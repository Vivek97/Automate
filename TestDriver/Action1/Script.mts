'*************************************************************************************************************************************************** @@ hightlight id_;_2296848_;_script infofile_;_ZIP::ssf77.xml_;_
'					Begin Test Driver
'***************************************************************************************************************************************************

Public fsolog
Public ofilelog
Public arrPublic()

strScriptPath = Environment.Value("TestDir")
strIndividualFolder = Split(strScriptPath,"\",-1,1)

strIntPath =""

' Getting Framework Path
For intCounter = 0 to UBound(strIndividualFolder) - 1
	strIntPath = strIntPath & strIndividualFolder(intCounter)  & "\"
	strIntPath = Trim(strIntPath)
Next
Environment.Value("FolderDirPath") = strIntPath
' Setting the path
strExcelPath = strIntPath & "DataSheet\" 
Environment.Value("ExcelPath") = strExcelPath
strLibPath = strIntPath & "FunctionLibrary\"
strResultPath = strIntPath & "Results\"
Environment.Value("Resource_Path") = strIntPath & "Object_Repository\"
Environment.Value("ResultPath") = strResultPath
'Added by Febin for Web services on 11th Mar 2013
Environment.Value("WSPath")  = strIntPath & "Resources\"
'Keeping Computer unlock in batch execution to avoid object identiifaction error failuer from windows auto lock
CreateObject("WScript.Shell").Run(strIntPath & "\Initialization\DONOTLOCK.vbs") 

'Import the Master Data File
DataTable.AddSheet("MasterSheet")
DataTable.ImportSheet Environment.Value("ExcelPath") & "MasterData.xls", "MasterSheet", "MasterSheet"
DataTable.AddSheet("GLOBALPARAMETERS")
DataTable.ImportSheet Environment.Value("ExcelPath") & "Parameters.xls", "GLOBALPARAMETERS", "GLOBALPARAMETERS"
Environment.Value("ScreenShotForPass") = DataTable.Value("GParam_ScreenshotForPass", "GLOBALPARAMETERS")

' Business Unit Excel File Path
Environment.Value("BUExcelPath") = Environment.Value("ExcelPath") & "BusinessUnits.xls"

DataTable.AddSheet("TestLab")
DataTable.AddSheet("LOCALPARAMETERS")
DataTable.AddSheet("ITERATIONPARAMETERS")
Environment.Value("MasterSheetRow") = 1
CreateExecutionSummaryFile()											' Create Excel Report
CreateLog("ExecutionLog")                                                  

While DataTable.Value("SCENARIO_NAME", "MasterSheet") <> "END"
		strExecuteTC = DataTable.Value("EXECUTE", "MasterSheet")
		DataTable.GetSheet("MasterSheet").GetCurrentRow	
		If  UCase(strExecuteTC) = "YES" Then
				Environment.Value("TestCaseFileName") = DataTable.Value("SCENARIO_FILENAME", "MasterSheet")
				strTCExcelPath = Environment.Value("ExcelPath") & Environment.Value("TestCaseFileName")					' Test Case file path
				Environment.Value("strTestCase") = DataTable.Value("SCENARIO_NAME", "MasterSheet")
				Call Execution_log(Environment.Value("strTestCase"), "", "", "") 'Create Log file with name ExecutionLog, it is a common log file for entire execution
                CreateHtmlReport()			' Create HTML Report
                Environment.Value("TESTSCENARIOCOUNT") = 0
				Environment.Value("TOTALFAILCOUNT") = 0
				Environment.Value("intFailCount") = 0
				Environment.Value("MasterSheetRow") = DataTable.GetSheet("MasterSheet").GetCurrentRow
				DataTable.ImportSheet strTCExcelPath,"TestLab", "TestLab"
				DataTable.ImportSheet strTCExcelPath,"LOCALPARAMETERS", "LOCALPARAMETERS"
				DataTable.ImportSheet strTCExcelPath,"ITERATIONPARAMETERS", "ITERATIONPARAMETERS"
				Environment.Value("TestLabRow") = 1
				DataTable.GetSheet("TestLab").SetCurrentRow(1)
				While DataTable.Value("TESTCASE_NAME","TestLab") <> "END" 	' Loop until the TESTCASE_NAME is END in the Scenario sheet
                        strExecute = DataTable.Value("EXECUTE","TestLab")
						DataTable.GetSheet("TestLab").GetCurrentRow			' Get the Test Scenario to be run
						If  UCase(strExecute) = "YES" Then
							Environment.Value("TESTSCENARIOCOUNT") = Environment.Value("TESTSCENARIOCOUNT") + 1
							Environment.Value("TestStepLog") = "True"
							strScenarioName = DataTable.Value("TESTCASE_NAME","TestLab")
							Environment.Value("ScenarioName") = strScenarioName
							Environment.Value("ExitOnFailure") = DataTable.Value("EXIT_TESTCASE_ONFAILURE","TestLab")
							Call ExecuteTC(strTCExcelPath, strScenarioName)					' Call the ExecuteTC to navigate through the Test Data file
							If  Environment.Value("TestStepLog") = "False" Then
									Environment.Value("TOTALFAILCOUNT") = Environment.Value("TOTALFAILCOUNT") + 1
									Call Execution_log("", Environment.Value("ScenarioName"), "", "Fail")
							Else
									Call Execution_log("", Environment.Value("ScenarioName"), "", "Pass")
							End If
						End If
						Wait(1)
						DataTable.SetCurrentRow(Environment.Value("TestLabRow"))
						DataTable.GetSheet("TestLab").SetNextRow
						Environment.Value("TestLabRow") = DataTable.GetSheet("TestLab").GetCurrentRow
				WEnd   
				CloseReport() 				' Closing the HTMLreport	
		End If
	DataTable.SetCurrentRow(Environment.Value("MasterSheetRow"))
    DataTable.GetSheet("MasterSheet").SetNextRow
	Environment.Value("MasterSheetRow") = DataTable.GetSheet("MasterSheet").GetCurrentRow
WEnd
CloseLog()						' closing the Log file	
CloseExecutionSummary()			' Closing the Excel report
'****************************************************************************************************************************************************
'                            						End Of Test Script
'****************************************************************************************************************************************************