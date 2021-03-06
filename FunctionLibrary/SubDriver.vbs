Public strDataVal
Public ParamIteration
Public intIteration
Public IterationObject
Public strIteFlag

strIteFlag="False"
strBUIteFlag="False"
Environment.Value("blnConditionFlag")="False"																								'if true Conditional flow is getting executed		
Environment.Value("ExecutingBU")="False"																										'If true BU is getting executed
Environment.Value("blnExecutingIteration")="False"																							'If true Iteration is getting executed	

'==================================================================================================================================================
' Name of the Function     			  : ExecuteTC()
' Description       		   		 	     : This function will navigate through the DataTable and make a call to the appropriate STEP (Action, BU, GUI_Function, NON-GUI_Function etc
' Date and / or Version       	    : 
' Example Call							 : ExecuteTC("TestScenario1")
'==================================================================================================================================================
Function ExecuteTC(strTCExcelPath, strScenarioName)

			StepStartTime =Time												'Added to keep track of time of execution for reporting part by Basavaraj on 15th April, 2013
			Environment.Value("BUFlag") = True
			Environment.Value("BUCall") = 0
			Environment.Value("SubBUCall") = 0
			Environment.Value("IterationParamRow") = 0
			Set ParamIteration = CreateObject("Scripting.Dictionary")				' Create a Dictionary object to count the no. of calls made to Local Param
			Set IterationObject = CreateObject("Scripting.dictionary")

			DataTable.AddSheet strScenarioName
			DataTable.ImportSheet strTCExcelPath, strScenarioName, strScenarioName

            Datatable.GetSheet(strScenarioName).SetCurrentRow(1)
			strFirstStep = Datatable.Value("STEP",strScenarioName)
			If  strFirstStep = "ITERATION" Then
					intIteration = Datatable.Value("INPUTDATA_PARAMETER",strScenarioName)
					intStartRow = 2
			Else
					intIteration = 1
					intStartRow = 1
			End If

			Environment.Value("ScenarioSheetRow") = intStartRow
			For intExecutionIteration = 1 to intIteration

						'*********************************************************************************************
						'Description: Store itteration value in environment variable to know the current itteration in BU
						'Author:      Vignesh
						'Date:	      Jan 20 2012
						'*********************************************************************************************
						Environment.Value("CurrIterationCount") = intExecutionIteration
						Environment.Value("IterationParamRow") = Environment.Value("IterationParamRow") + 1
						'*********************************************************
						'Sets the TestObjectFlag to true for each itteration
						'Vignesh
						'20 Jan 2012
						'*********************************************************
						Environment.Value("TestObjectFlag") = "True"
						Environment.Value("flgrExitItration") = "False"
	
						If  intIteration > 1 Then
							UpdateReport "TESTCASE", strScenarioName & " - Iteration " &intExecutionIteration , "", "", "", "", ""
						ElseIf intIteration = 1 Then
							UpdateReport "TESTCASE", strScenarioName, "", "", "", "", ""
						End If		
						
						Datatable.GetSheet(strScenarioName).SetCurrentRow(intStartRow)
						'*********************************************************
						'While....Wend loop changed to Do While.....Loop
						'Vignesh
						'27 Dec 2011 
						'*********************************************************
						Do While DataTable.Value("STEP",strScenarioName) <> "END" 

							'On Error Resume Next
							Datatable.GetSheet(strScenarioName).GetCurrentRow
							strDataVal = Datatable.Value("INPUTDATA_PARAMETER", strScenarioName)
							strStep = Datatable.Value("STEP",strScenarioName)
							Call ExecuteDataManage(strScenarioName,strDataVal,strStep) 
	
								If  (Environment.Value("TestObjectFlag") = "False" OR Environment.Value("TestStepLog") = "False") AND UCASE(Environment.Value("ExitOnFailure")) = "YES" Then
	
										UpdateReport "TESTSTEP", "", "<font color=""red"">Test Execution Status</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Test Execution for this scenario is Stopped due to the failure in the above step</font>", "Stop"
	
										'**************************Added by Anjan on 12/03/2015 to Close MCS Power builder applications for Test Object Flag************************
										Call fn_CloseAllPBApplications(strScenarioName,strDataVal)
										'*************************************************************************************************************************************************************
										Call fn_CloseAllBrowser()
										'************************************************************************************
										'Added on 27 Dec 2011 
										'Author Vignesh Somasekar
										'Reason: Removes all the dictionary object if any fail in the test itteration
										'					and also exit of the current itteration and continue with next itteration
										'************************************************************************************
										If Environment.Value("flgrExitItration") = "True" Then
											ParamIteration.RemoveAll
											IterationObject.RemoveAll
											Environment.Value("flgrExitItration") = "False"
											Exit Do
										Else
											'********************************
											'For backward compatibility
											'********************************
											'Exit Function			' Exit the loop to upload the result file into QC(To Be Decided)
	
											Exit Do			' Exit the loop to upload the result file into QC
										End If
								End If

								DataTable.GetSheet(strScenarioName).SetCurrentRow(Environment.Value("ScenarioSheetRow"))
								DataTable.GetSheet(strScenarioName).SetNextRow
								Environment.Value("ScenarioSheetRow") = DataTable.GetSheet(strScenarioName).GetCurrentRow
								Environment.Value("ScenarioSheetBUrow")  =  DataTable.GetSheet(strScenarioName).GetCurrentRow

						Loop

				Environment.Value("ScenarioSheetRow") = intStartRow
				'********************************
				'For backward compatibility
				'********************************
				If Environment.Value("flgrExitItration") = "True" Then
					ParamIteration.RemoveAll
					IterationObject.RemoveAll
				End If

			Next																																							'Loop Next for Iteration Controller
			'DataTable.DeleteSheet(strScenarioName) 
End Function


'==================================================================================================================================================
' Name of the Function     			  : ExecuteBU
' Description       		   		 	      : This function will navigate through the DataTable and make a call to the appropriate STEP (Action, BU, GUI_Function, NON-GUI_Function etc
' Date and / or Version       	    : 
' Example Call							  : ExecuteBU(LocalParamIteration, GlobalParamIteration,"TIPS_LOGIN")
'==================================================================================================================================================

Function ExecuteBU(ParamIteration, strBUName)

		Dim strDataCount(100)
    	Environment.Value(strBUName) = 1
		Environment.Value("BUSheetRow") = Environment.Value(strBUName)

		While DataTable.Value("STEP",strBUName) <> "END"			' Loop until the STEP is END in the BU sheet
				Datatable.GetSheet(strBUName).GetCurrentRow
				strBUStep = Datatable.Value("STEP", strBUName)
				If  strBUStep = "BUSINESS_UNIT" Then
					Environment.Value("BUFlag") = False
				End If

				strDataVal = Datatable.Value("INPUTDATA_PARAMETER", strBUName)
				Call ExecuteDataManage(strBUName,strDataVal,strBUStep)

				If  (Environment.Value("TestObjectFlag") = "False" OR Environment.Value("TestStepLog") = "False") AND UCASE(Environment.Value("ExitOnFailure")) = "YES" Then
						Exit Function
				End If

				DataTable.GetSheet(strBUName).SetCurrentRow(Environment.Value("BUSheetRow"))
				DataTable.GetSheet(strBUName).SetNextRow
				Environment.Value("BUSheetRow") = DataTable.GetSheet(strBUName).GetCurrentRow
		WEnd  

        UpdateLog "Exiting from the Business Unit"
		Environment.Value("ExecutingBU")="False"
'		DataTable.DeleteSheet(strBUName)  ' Added by shrinidhi on 10/16/12 to delete the Used BU sheet
End Function

'==================================================================================================================================================
' Name of the Function     			  : ExecuteAction
' Description       		   		 	      : This function will enter or retrieve the Object values based upon the type of Action
' Date and / or Version       	    : 
' Example Call							  : ExecuteAction("JOHN","TestScenario1")
'==================================================================================================================================================

Function ExecuteAction(strDataVal, strScenarioName)
		   strData = strDataVal
		   strAction_Name = Datatable.Value("ACTION_KEYWORD", strScenarioName)				' ACTION _KEYWORD value
		   strObjectHier = Datatable.value("APP_SCREEN_NAME", strScenarioName)		' Object Hierarchy value ex: Browser("Login").Page("Login")
		   strObject = Datatable.Value("OBJECT", strScenarioName)
		   Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)		
''*************************************************Added By Suresh on 02/08/2011 *********************************************************
'If strObject="" Then
'	arrActObject=Split(strObjectHier,".")
'	strObjectGet = Split(arrActObject(2), "(", -1)
'	strObjectNew = Split(strObjectGet(1), ")")
'	strObjectType = strObjectGet(0)																								' Object Type ex: WebEdit
'	strObjectVal = Split(strObjectNew(0),"""")
'	strObjectName = strObjectVal(1)	
'Else
'	strObjectGet = Split(strObject, "(", -1)
'	strObjectNew = Split(strObjectGet(1), ")")
'	strObjectType = strObjectGet(0)																								' Object Type ex: WebEdit
'	strObjectVal = Split(strObjectNew(0),"""")
'	strObjectName = strObjectVal(1)
'End If

'*******************************************************************************************************************************************
' Added by Suresh, modified by Shrinidhi  - Aug 2nd 2011, modified by Febin for Terminal Emulator on Feb 21-2012
'*******************************************************************************************************************************************
			If  strObject = "" And InStr(UCASE(strObjectHier), "WEBTABLE(") > 0 Then						' Generally strObject will be blank for WebTables, the hierarchy will be Browser().Page().Frame().WebTable()
				strTableObjSplit = Split(strObjectHier, ".")
				intUBound = UBound(strTableObjSplit)
				strTableObj = strTableObjSplit(intUBound)				' Always the last object in the hierarchy will be the Tableobject
				strObjectGet = Split(strTableObj, "(", -1)
			ElseIf strObject = "" And InStr(UCASE(strObjectHier), "TEScreen(") = 0  Then ' Identifying the Terminal Emulator object where there won't be any browser
                strObjectGet = Split(strObjectHier, "(", -1)
			Else
				strObjectGet = Split(strObject, "(", -1)
			End If
' ********************************************************************************************************************************************
'		   strObjectGet = Split(strObject, "(", -1)
		   strObjectNew = Split(strObjectGet(1), ")")
		   strObjectType = strObjectGet(0)																								' Object Type ex: WebEdit
		   strObjectVal = Split(strObjectNew(0),"""")
		   strObjectName = strObjectVal(1)																							  ' Object Name ex: LoginName
		   Select Case strObjectType
				   Case "WebEdit" 												' Call to the appropriate Object Type
					  Call OperateOnWebEdit(strObjectHier, strObjectName, strAction_Name, strData)
				
				   Case "WebButton" 
					  Call OperateOnWebButton(strObjectHier, strObjectName, strAction_Name, strData)
				
				   Case "WebList" 
					  Call OperateOnWebList(strObjectHier, strObjectName, strAction_Name, strData)
				
				   Case "Link" 
					  Call OperateOnLink(strObjectHier, strObjectName, strAction_Name, strData)
				
				   Case "WebRadioGroup" 
					  Call OperateOnWebRadioGroup(strObjectHier, strObjectName, strAction_Name, strData)
				
				   Case "WebElement" 
					  Call OperateOnWebElement(strObjectHier, strObjectName, strAction_Name, strData)
					  
                   Case "Image" 
					  Call OperateOnImage(strObjectHier, strObjectName, strAction_Name, strData)

				   Case "WebCheckBox" 
					  Call OperateOnWebCheckBox(strObjectHier, strObjectName, strAction_Name, strData) 

				    Case "WinEdit" 												' Call to the appropriate Object Type
					  Call OperateOnWinEdit(strObjectHier, strObjectName, strAction_Name, strData)
					  
					Case "WinEditor" 												' Call to the appropriate Object Type
					  Call OperateOnWinEditor(strObjectHier, strObjectName, strAction_Name, strData)

				    Case "WinButton" 												' Call to the appropriate Object Type
					  Call OperateOnWinButton(strObjectHier, strObjectName, strAction_Name, strData)

				    Case "WinComboBox" 												' Call to the appropriate Object Type
					  Call OperateOnWinComboBox(strObjectHier, strObjectName, strAction_Name, strData)

				    Case "WinCheckBox" 												' Call to the appropriate Object Type
					  Call OperateOnWinCheckBox(strObjectHier, strObjectName, strAction_Name, strData)

					Case "WinRadioButton" 												' Call to the appropriate Object Type
					  Call OperateOnWinRadioButton(strObjectHier, strObjectName, strAction_Name, strData)

					Case "WinMenu" 												' Call to the appropriate Object Type
					  Call OperateOnWinMenu(strObjectHier, strObjectName, strAction_Name, strData)
					  
				  	Case "WinList" 												' Call to the appropriate Object Type
					  Call OperateOnWinList(strObjectHier, strObjectName, strAction_Name, strData)
                					
					Case "PbList" 												' Call to the operate on Power Builber List Object Type
					  Call OperateOnPbList(strObjectHier, strObjectName, strAction_Name, strData)
					
					Case "PbEdit" 												' Call to the Operate on Power builder Edit Object Type
					  Call OperateOnPbEdit(strObjectHier, strObjectName, strAction_Name, strData)
					 
					Case "PbDataWindow" 												' Call to the Operate on Power builder Data Window Object Type
					  Call OperateOnPbDataWindow(strObjectHier, strObjectName, strAction_Name, strData)					 
					
					Case "PbButton" 												' Call to the Operate on Power builder Button Object Type
					  Call OperateOnPbButton(strObjectHier, strObjectName, strAction_Name, strData)
					
					Case "PbObject"
					  Call OperateOnPbObject(strObjectHier, strObjectName, strAction_Name, strData)	
					
					Case "Static"													'Call to Operate on Power Builber Static Message which appears on DialgoBox Object Type
					  Call OperateOnStaticMessage(strObjectHier, strObjectName, strAction_Name, strData) 
					
					Case "PbComboBox"												 'Call to Operate on Power Builber ComboBox Object Type
					  Call OperateOnPbComboBox(strObjectHier, strObjectName, strAction_Name, strData) 
  
					Case "PbCheckBox"												 'Call to Operate on Power Builber CheckBox Object Type
					  Call OperateOnPbCheckBox(strObjectHier, strObjectName, strAction_Name, strData) 
  
					
					Case "PbRadioButton" 												' Call to the Operate on Power builder RadioButton Object Type
					  Call OperateOnPbRadioButton(strObjectHier, strObjectName, strAction_Name, strData)
					
					Case "SAPGuiEdit"
						 Call OperateOnSAPEdit(strObjectHier, strObjectName, strAction_Name, strData)
						 
				   Case  "SAPGuiButton"
					     Call OperateOnSAPButton(strObjectHier, strObjectName, strAction_Name, strData)
						 
				   Case "SAPGuiRadioButton"
					      Call OperateOnSAPRadioGroup(strObjectHier, strObjectName, strAction_Name, strData)
						  
				   Case "SAPGuiTabStrip"
					      Call OperateOnSAPTabStrip(strObjectHier, strObjectName, strAction_Name, strData)
						  
				   Case "SAPGuiTable"
						  Call OperateOnSAPGUITable(strObjectHier, strObjectName, strAction_Name, strData)
						  
				   Case "SAPGuiOKCode"
						Call OperateOnSAPGUIOKCode(strObjectHier, strObjectName, strAction_Name, strData)
						
				   Case "SAPGuiStatusBar"		
						Call ReadSAPStatusbar(strObjectHier, strObjectName, strAction_Name, strData)
						
				   Case "SAPGuiGrid"	
						Call OperateOnSAPGrid(strObjectHier, strObjectName, strAction_Name, strData)
						
				   Case "SAPGuiTree"	
						Call OperateOnSAPTree(strObjectHier, strObjectName, strAction_Name, strData)
'**************************************************Added by Suresh on 02/08/2011 *********************************************************					
		           Case "WebTable"
						Call OperateOnWebTable(strObjectHier, strAction_Name, strData)

					 Case "WebFile"
						Call OperateOnWebFile(strObjectHier, strObjectName, strAction_Name, strData)
'**************************************************************************************************************************************
'************************************************* Added by Febin Mathew on Feb 21-2012*************************************
				   Case "TEField"
						strSplitObjectHier = Split(strObjectHier, "(")
						strObjectHier = Mid(strSplitObjectHier(1), 2,Len(strSplitObjectHier(1)) - 3)	' Getting the Screen Name followed by the description and taking out "TEScreen"
						Call OperateOnTerminalEmulator(strObjectHier, strObjectName, strAction_Name, strData)
												
				   Case "TEScreen"
                        Call OperateOnTerminalEmulator(strObjectName, "", strAction_Name, strData)		' In this case, the strObjectName is nothing but the hierarchy name. And 2nd parameter is not needed hence passing it as blank
'***************************************************************************************************************************************
		   End Select
End Function  


'==================================================================================================================================================
' Name of the Function     			  : ExecuteIteration()
' Description       		   		 	     : This function will navigate through the DataTable and make a call to the appropriate STEP (Action, BU, GUI_Function, NON-GUI_Function etc
' Date and / or Version       	    : 
' Example Call							 : ExecuteIteration("TestScenario1")
'==================================================================================================================================================
Function ExecuteIteration(ParamIteration, strScenarioName)

			Dim strDataCount(100)
			Environment.Value("SubIterationCall") = 0
           DataTable.GetSheet(strScenarioName).SetCurrentRow(Environment.Value("ScenarioSheetRow"))
		   	intSubIteration = Datatable.Value("INPUTDATA_PARAMETER",strScenarioName)

			If Instr(intSubIteration,"fn_") > 0 or Instr(intSubIteration,"VAR_") > 0 Then
				intSubIteration = Environment.Value(intSubIteration)
			End If

		    DataTable.GetSheet(strScenarioName).SetNextRow
			Environment.Value("ScenarioSheetIterationRow") = DataTable.GetSheet(strScenarioName).GetCurrentRow
			For intSubExecutionIteration = 1 to intSubIteration
					Environment.Value("SubCurrIterationCount") = intSubExecutionIteration
					DataTable.GetSheet(strScenarioName).SetCurrentRow(Environment.Value("ScenarioSheetIterationRow"))
					Environment.Value("blnExecutingIteration")="True"
					Environment.Value("TestObjectFlag") = "True"
					Environment.Value("flgrExitItration") = "False"

					UpdateReport "TESTCASE", strScenarioName & " - Sub Iteration " &intSubExecutionIteration , "", "", "", "", ""
					Do While DataTable.Value("STEP",strScenarioName) <> "END_ITERATION"

						'On Error Resume Next
							Datatable.GetSheet(strScenarioName).GetCurrentRow
							strStep = Datatable.Value("STEP",strScenarioName)
							strDataVal = Datatable.Value("INPUTDATA_PARAMETER", strScenarioName)
							Call ExecuteDataManage(strScenarioName,strDataVal,strStep)

							If  (Environment.Value("TestObjectFlag") = "False" OR Environment.Value("TestStepLog") = "False") AND UCASE(Environment.Value("ExitOnFailure")) = "YES" Then

									UpdateReport "TESTSTEP", "", "<font color=""red"">Test Execution Status</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Test Execution for this scenario is Stopped due to the failure in the above step</font>", "Stop"
									'**************************Added by Febin on 2/28/2012 for Closing TErminal Emulator sessions for Test Object Flag************************
									Call fn_CloseAllPBApplications(strScenarioName,strDataVal)
									'*************************************************************************************************************************************************************
									Call fn_CloseAllBrowser()
									'************************************************************************************
									'Added on 27 Dec 2011 
									'Author Vignesh Somasekar
									'Reason: Removes all the dictionary object if any fail in the test itteration
									'					and also exit of the current itteration and continue with next itteration
									'************************************************************************************
									If Environment.Value("flgrExitItration") = "True" Then
										ParamIteration.RemoveAll
										Environment.Value("flgrExitItration") = "False"
										Exit Do
									Else
										'********************************
										'For backward compatibility
										'********************************
										'Exit Function		'To Be decided
										Exit Do			' Exit the loop to upload the result file into QC
									End If
							End If
							DataTable.GetSheet(strScenarioName).SetNextRow
							Environment.Value("ScenarioSheetRow") = DataTable.GetSheet(strScenarioName).GetCurrentRow
					Loop	'
					UpdateReport "TESTCASE", strScenarioName & " -End of  Sub Iteration " &intSubExecutionIteration , "", "", "", "", ""
			Next
			'Environment.Value("ScenarioSheetRow") = DataTable.GetSheet(strScenarioName).GetNextRow
End Function


'==================================================================================================================================================
' Name of the Function     			  : ExecuteDataManage()
' Description       		   		 	     : This function will manage the test data
' Date and / or Version       	    : 
' Example Call							 : ExecuteDataManage("TestScenario1")
'==================================================================================================================================================
Function ExecuteDataManage(strScenarioName,strDataVal,strStep)

							Dim strDataCount(100)							
							strDataValSplit = Split(strDataVal, ",")
							strSequence=DataTable.Value("SEQUENCE",strScenarioName)
							'Set ParamIteration = CreateObject("Scripting.Dictionary")	'To be decided
							If  UBOUND(strDataValSplit) > 0 Then													' When Input Parameter contains more than one value ex: GParam_CaseID, LParam_LoanNumber, ABC
									For intCount = 0 to UBOUND(strDataValSplit)
											If Instr(Trim(strDataValSplit(intCount)), "LParam_") > 0 Then					' If Input Parameter starts with Param_, retrieve the value from local Parameter sheet
													strParamName = Trim(strDataValSplit(intCount))
													If ParamIteration.Exists(strParamName) Then														' If Param_ Name already exists in the Dictionary, append the count by 1
															Environment.Value(strParamName) = Environment.Value(strParamName) + 1
													Else
															ParamIteration.Add strParamName, strStep													' If the Param_ Name does not exist in the Dictionary, create an entry and update the Call count to 1
															Environment.Value(strParamName) = 1
													End If

													DataTable.GetSheet("LOCALPARAMETERS").SetCurrentRow(Environment.Value("TestLabRow"))			' Setting the Current row equal to TestCase Row in the Test Case sheet
													strParamDataValue = DataTable.Value(strParamName,"LOCALPARAMETERS")		
													
													'***********************************************************
													'Added on 27 Dec 2011 
													'Author Vignesh Somasekar
													'Reason: to handle LParam in case of itteration exit
													'***********************************************************
													strIterVar	= strParamDataValue
													If intIteration > 1 and Instr(1, strIterVar, "||") > 0 Then
														strarrItrVars = Split(strIterVar, "||")
														strParamDataValue = strarrItrVars(Environment.Value("CurrIterationCount")-1)
														Environment.Value("flgrExitItration") = "True"
													End If
													'***********************************************************
															
													strDataSplit = Split(strParamDataValue, "^^")																					' Split the Data value by ^^ separator
													If UBound(strDataSplit) = 0 Then
															strDataVal = strParamDataValue																			' When the LParam do not contain ^^ separator
													ElseIf UBound(strDataSplit) = -1 Then
															strDataVal = ""													' When the LParam contain ^^ separator and data is intentionally kept blank, the UBound will be -1. Code added by Shrinidhi on Aug 01 2011.
													Else
															strDataVal = strDataSplit(Environment.Value(strParamName)-1)
													End If

											ElseIf Instr(Trim(strDataValSplit(intCount)), "IParam_") > 0 Then					' If Input Parameter starts with IParam_, retrieve the value from Iteration Parameter sheet
													strParamName = Trim(strDataValSplit(intCount))
													If ParamIteration.Exists(strParamName) Then														' If Param_ Name already exists in the Dictionary, append the count by 1
														ParamIteration.Item(strParamName) = ParamIteration.Item(strParamName) + 1
															Environment.Value(strParamName) = Environment.Value(strParamName) + 1
													Else
															'ParamIteration.Add strParamName, 0													' If the Param_ Name does not exist in the Dictionary, create an entry and update the Call count to 1
															ParamIteration.Add strParamName, 0
															Environment.Value(strParamName) = 1
													End If

													Call fn_FindIterationParamRowNumber(Environment.Value("ScenarioName"))
													Environment.Value("IterationParamRow") = Environment.Value("IterationParamRow") + ParamIteration.Item(strParamName)
													DataTable.GetSheet("ITERATIONPARAMETERS").SetCurrentRow(Environment.Value("IterationParamRow"))
													'Environment.Value("IterationParamRow") = Environment.Value("IterationParamRow") + 1										' Setting the Current row equal to TestCase Row in the Test Case sheet  based on number of calls made on particular param
 													strParamDataValue = DataTable.Value(strParamName,"ITERATIONPARAMETERS")	
                                                    strDataVal = strParamDataValue

											ElseIf Instr(Trim(strDataValSplit(intCount)), "GParam_") > 0 AND Datatable.Value("ACTION_NAME", Environment.Value("ScenarioName")) <> "fn_UpdateValueGlobalParameter" Then					' If Input Parameter starts with Param_, retrieve the value from Global Parameter sheet and the Keyowrd is not fn_UpdateValueGlobalParameter function
													strParamName = Trim(strDataValSplit(intCount))
													DataTable.GetSheet("GLOBALPARAMETERS").SetCurrentRow(1)			' Setting the Current row equal to 1
													strDataVal = DataTable.Value(strParamName,"GLOBALPARAMETERS")
											ElseIf Instr(Trim(strDataValSplit(intCount)), "fn_") = 1 OR Instr(Trim(strDataValSplit(intCount)), "VAR_") = 1 Then		'Added by Shrinidhi on 5th mar 2012 - If the data is a return value from a function or another gloabl variable
													strDataVal = Environment.Value(strDataValSplit(intCount))
											Else
													strDataVal = Trim(strDataValSplit(intCount))											' When the data is not LParam or GParam
											End If
											strDataCount(intCount) = strDataVal
											If  intCount = 0 Then
													strDataValFinal = strDataCount(intCount)																' 
											Else
													strDataValFinal = strDataValFinal & "," & strDataCount(intCount)
											End If
									Next
									strDataVal = strDataValFinal
							Else																						' When only one value is mentioned, ex: GParam_URL
									If Instr(strDataVal, "LParam_") > 0 Then					' If Input Parameter starts with Param_, retrieve the value from local Parameter sheet
											strParamName = strDataVal
											If ParamIteration.Exists(strParamName) Then														' If Param_ Name already exists in the Dictionary, append the count by 1
													Environment.Value(strParamName) = Environment.Value(strParamName) + 1
											Else
													ParamIteration.Add strParamName, strStep													' If the Param_ Name does not exist in the Dictionary, create an entry and update the Call count to 1
													Environment.Value(strParamName) = 1
											End If
											'Call fn_FindLocalParamRowNumber(Environment.Value("ScenarioName"))
											DataTable.GetSheet("LOCALPARAMETERS").SetCurrentRow(Environment.Value("TestLabRow"))			' Setting the Current row equal to TestCase Row in the Test Case sheet
											strParamDataValue = DataTable.Value(strParamName,"LOCALPARAMETERS")	

											'***********************************************************
											'Added on 27 Dec 2011 
											'Author Vignesh Somasekar
											'Reason: to handle LParam in case of itteration exit
											'***********************************************************
											strIterVar	= strParamDataValue
											If intIteration > 1 and Instr(1, strIterVar, "||") > 0 Then
												strarrItrVars = Split(strIterVar, "||")
												strParamDataValue = strarrItrVars(Environment.Value("CurrIterationCount")-1)
												Environment.Value("flgrExitItration") = "True"
											End If
											'***********************************************************
																
											strDataSplit = Split(strParamDataValue, "^^")																					' Split the Data value by ^^ separator
											If UBound(strDataSplit) = 0 Then
													strDataVal = strParamDataValue																			' When the LParam do not contain ^^ separator
											ElseIf UBound(strDataSplit) = -1 Then
															strDataVal = ""													' When the data is intentionally kept blank
											Else
													strDataVal = strDataSplit(Environment.Value(strParamName)-1)											  ' This Data value depends on the no. of times Param_ is used. For ex: if Param_ABC  is used 2nd time then 2nd parameter under "Param_ABC" in PARAMETERS Sheet column will be picked 
											End If

									ElseIf Instr(strDataVal, "IParam_") > 0 Then					' If Input Parameter starts with Param_, retrieve the value from Iteration Parameter sheet
													strParamName = Trim(strDataValSplit(intCount))
													If ParamIteration.Exists(strParamName) Then														' If Param_ Name already exists in the Dictionary, append the count by 1
														ParamIteration.Item(strParamName) = ParamIteration.Item(strParamName) + 1
'															Environment.Value(strParamName) = Environment.Value(strParamName) + 1
													Else
															ParamIteration.Add strParamName, 0													' If the Param_ Name does not exist in the Dictionary, create an entry and update the Call count to 1
'															Environment.Value(strParamName) = 1
													End If
													
													Call fn_FindIterationParamRowNumber(Environment.Value("ScenarioName"))
													Environment.Value("IterationParamRow") = Environment.Value("IterationParamRow") + ParamIteration.Item(strParamName)
													DataTable.GetSheet("ITERATIONPARAMETERS").SetCurrentRow(Environment.Value("IterationParamRow"))			' Setting the Current row equal to TestCase Row in the Test Case sheet based on number of calls made on param
													'Environment.Value("IterationParamRow") = Environment.Value("IterationParamRow") + 1
													strParamDataValue = DataTable.Value(strParamName,"ITERATIONPARAMETERS")	
                                                    strDataVal = strParamDataValue

									ElseIf Instr(strDataVal, "GParam_") > 0 AND Datatable.Value("ACTION_NAME", Environment.Value("ScenarioName")) <> "fn_UpdateValueGlobalParameter" Then					' If Input Parameter starts with Param_, retrieve the value from Global Parameter sheet and the Keyowrd is not fn_UpdateValueGlobalParameter function
											strParamName = strDataVal
											DataTable.GetSheet("GLOBALPARAMETERS").SetCurrentRow(1)			' Setting the Current row equal to 1
											strDataVal = DataTable.Value(strParamName,"GLOBALPARAMETERS")	
									ElseIf Instr(Trim(strDataVal), "fn_") = 1 OR Instr(Trim(strDataVal), "VAR_") = 1 Then		'Added by Shrinidhi on 5th Mar 2012 - If the data is a return value from a function or another gloabl variable
													strDataVal = Environment.Value(strDataVal)			
									End If
							End If								
							If Instr(strDataVal, "DB_EXCLUDED") = 0  Then					' Added by Shrinidhi 5th Mar 2012, if the current step do not require a return value from DB when DB is not connected intentionally
								Select Case strStep
										Case "ACTION"
												UpdateLog "Executing the step - """ & strStep & """, Sequence ID: " & Datatable.Value("SEQUENCE", strScenarioName) & ", Step Description: " & Datatable.Value("STEP_DESCRIPTION", strScenarioName) & ", Keyword: " & Datatable.Value("ACTION_NAME", strScenarioName) & ", Input Data used: " & strDataVal                                        ' Call Update Log function to log the step details for ACTION
												Call ExecuteAction(strDataVal,strScenarioName)
										Case "GUI_FUNCTION" 
												UpdateLog "Executing the step - """ & strStep & """, Sequence ID: " & Datatable.Value("SEQUENCE", strScenarioName) & ", Step Description: " & Datatable.Value("STEP_DESCRIPTION", strScenarioName) & ", Function Name: " & Datatable.Value("ACTION_NAME", strScenarioName) & ", Input Data used: " & strDataVal                       ' Call Update Log function to log the step details for GUIFunction
												Execute "Call " & Datatable.Value("ACTION_NAME", strScenarioName) & "(""" & strScenarioName & """, """ & strDataVal & """)"
										Case "NON-GUI_FUNCTION" 
												UpdateLog "Executing the step - """ & strStep & """, Sequence ID: " & Datatable.Value("SEQUENCE", strScenarioName) & ", Step Description: " & Datatable.Value("STEP_DESCRIPTION", strScenarioName) & ", Function Name: " & Datatable.Value("ACTION_NAME", strScenarioName) & ", Input Data used: " & strDataVal                       ' Call Update Log function to log the step details for NONGUIFunction
												Execute "Call " & Datatable.Value("ACTION_NAME", strScenarioName) & "(""" & strDataVal & """)"
										Case "BUSINESS_UNIT" 
												If Environment.Value("BUFlag") = True Then			
														UpdateLog "Executing the step - """ & strStep & """, Sequence ID: " & Datatable.Value("SEQUENCE", strScenarioName) & ", Step Description: " & Datatable.Value("STEP_DESCRIPTION", strScenarioName) & ", Business Unit Name: " & Datatable.Value("ACTION_NAME", strScenarioName)                                                                              ' Call Update Log function to log the step details for Business Unit
														strBUSheetName = Datatable.Value("ACTION_NAME", strScenarioName)
														Environment.Value("ParentBU") = strBUSheetName						' Setting the BU Name in Global variable for future reference.
														On Error Resume Next
																DataTable.GetSheet(strBUSheetName)				' Checking whether a sheet with the same name exists
																If Err.Number = 0 Then
																		strBUSheetName_New = strBUSheetName & "_BU_" & Environment.Value("BUCall")		' If sheet already exists, add a new name to the sheet
																Else
																		strBUSheetName_New = strBUSheetName				' Else same BU Name will be kept
																End If
'														On Error GoTo 0
														DataTable.AddSheet strBUSheetName_New
														DataTable.ImportSheet Environment.Value("BUExcelPath"), strBUSheetName, strBUSheetName_New
														Datatable.GetSheet(strBUSheetName_New).SetCurrentRow(1)				' Setting the DataTable current row of the BU Sheet to 1 
														Environment.Value("BUCall") = Environment.Value("BUCall") + 1
														Call ExecuteBU(ParamIteration, strBUSheetName_New)
												Else
														UpdateLog "Executing the step - """ & strBUStep & """, Sequence ID: " & Datatable.Value("SEQUENCE", strBUName) & ", Step Description: " & Datatable.Value("STEP_DESCRIPTION", strBUName) & ", Business Unit Name: " & Datatable.Value("ACTION_NAME", strBUName)
														strSubBUSheetName = Datatable.Value("ACTION_NAME", strBUName)
														If strSubBUSheetName = Environment.Value("ParentBU") Then											' If the called BU name is same as the parent BU Name, exit the function
															Reporter.ReportEvent micFail, "Invalid Call to BU Name", "BU Name is same as the Parent BU Name"
															Exit Function
														End If
														'On Error Resume Next
																DataTable.GetSheet(strSubBUSheetName)
																If Err.Number = 0 Then
																		strSubBUSheetName_New = strSubBUSheetName & "_SubBU_" & Environment.Value("SubBUCall")
																Else
																		strSubBUSheetName_New = strSubBUSheetName
																End If
														On Error GoTo 0
														DataTable.AddSheet strSubBUSheetName_New
														DataTable.ImportSheet Environment.Value("BUExcelPath"), strSubBUSheetName, strSubBUSheetName_New
														Datatable.GetSheet(strSubBUSheetName_New).SetCurrentRow(1)
														Environment.Value("SubBUCall") = Environment.Value("SubBUCall") + 1                                        
														'***********************************************************
														'Added on 13  March  2012 
														'Author Manish Kumar Singh
														'Reason: to Call Sub BU from Parent BU
														'***********************************************************
														Redim CollOfLastBuRowcount(Environment.Value("SubBUCall"))
														CollOfLastBuRowcount(Environment.Value("SubBUCall")) = Environment.Value("BUSheetRow")
						'								Call ExecuteBU(strSubBUExcelPath, LocalParamIteration, GlobalParamIteration, strSubBUSheetName_New)
														Call ExecuteBU(ParamIteration, strSubBUSheetName_New)		
														intBuIteration= Ubound(CollOfLastBuRowcount) 
																Environment.Value("BUSheetRow") = CollOfLastBuRowcount(intBuIteration)
														intBuIteration = intBuIteration -1
														'***********************************************************    																		
												End IF
										Case "ITERATION" 												
												UpdateLog "Executing the step - """ & strStep & """, Sequence ID: " & Datatable.Value("SEQUENCE", strScenarioName) & ", Step Description: " & Datatable.Value("STEP_DESCRIPTION", strScenarioName) & ", Function Name: " & Datatable.Value("ACTION_NAME", strScenarioName) & ", Input Data used: " & strDataVal                       ' Call Update Log function to log the step details for NONGUIFunction
												Call ExecuteIteration( ParamIteration, strScenarioName)
								End Select
							End If									
End Function

'==================================================================================================================================================
' Name of the Function                                                    : fn_FindIterationParamRowNumber
' Description                                                                                            : This function is used  to find the row number of the TC in Iteration Parameter sheet
' Date and / or Version                       : 
' Example Call                                                                                                     : 
'==================================================================================================================================================
Function fn_FindIterationParamRowNumber(strTestCaseName)
                                DataTable.GetSheet("ITERATIONPARAMETERS").SetCurrentRow(1)
                                For intCounter = 1 to DataTable.GetSheet("ITERATIONPARAMETERS").GetRowCount
                                                                If  DataTable.Value("TESTCASE_NAME", "ITERATIONPARAMETERS") = strTestCaseName  and Environment.Value("IterationParamRow") = Environment.Value("CurrIterationCount") Then
                                                                                                Environment.Value("IterationParamRow") = intCounter
                                                                                                Exit For
                                                                Else
                                                                                                DataTable.GetSheet("ITERATIONPARAMETERS").SetNextRow
                                                                End If
                                Next
End Function