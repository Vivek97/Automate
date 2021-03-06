'==================================================================================================================================================
' Name of the Function     			  : OperateOnPbObject
' Name of the Function     			  : OperateOnPbObject
' Description       		   		  : This function is used to enter values in to PbEdit object on Power Bulider Application
' Created by                          : Madhusudhana K S    
' Date and / or Version       	      : 20/30/2015
' Example Call						  : OperateOnPbObject (PbWindow("").PbWindow("").PbObject("dp_startdate").Type "")
'==================================================================================================================================================

Function OperateOnPbObject(strObjectHierarchy, strObjectName, strAction_Name, strData)
   
		If  strAction_Name <> "STOREVALUE" AND Instr(strData, "VAR_") = 1 Then				' If the value to be taken from already saved variable
			strData = Environment.Value(strData)
		End If

		If  Instr(strData, "fn_") = 1 Then
			strData = Environment.Value(strData)
		End If

		Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.PbObject(strObjectName)
		StepStartTime = Time
        If fnc_wait(ActualObject) = "True" Then
				Select Case strAction_Name
					  Case "TYPEVALUE"	
	  						  Wait(1)
	  						  
	  						  arrDataVal = Split(strData,";")
	  						  If UBound(arrDataVal) > 1 Then
	  						  	strData = arrDataVal(2)
	  						  End If
	  						  Position=Split(arrDataVal(0),",")
	  						  ObjectHierarchy.Activate
	  						  ActualObject.Click Position(0),Position(1)
							  ActualObject.Type arrDataVal(1)
							  strActualVal = ActualObject.GetROProperty("text")
							  
							  If StrComp(strActualVal,arrDataVal(1)) = 0  Then					   		
							  		UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime, Time, "Value: <font color=""blue""> <b><i>" & arrDataVal(1) & "</i></b></font> is entered " & " in <b>'" & Mid(strObjectName,4,Len(strObjectName)-3) & "'</b> field successfully", "Done"
							  Else
							  		UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "<font color=""blue"">The Value: <b><i>" & arrDataVal(1) & "</i></b> is not entered in field <b>" & strobjName &"</b></font>", "Fail"
							  		Environment.Value("TestStepLog") = "False"
							  End If		
							 			  
'					
						End Select
        Else
                UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The WebEdit - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If

End Function

'==================================================================================================================================================
' Name of the Function     			: OperateOnStaticMessage
' Description       		   		: This function is used to operate on static message in a dialogbox
' Author 							: Akshatha
'Date and / or Version       	    : 26/03/2015 
' Example Call						: OperateOnStaticMessage("PbWindow("abc")","OK", "Set", "")
'==================================================================================================================================================

Function OperateOnStaticMessage(strObjectHierarchy, strObjectName, strAction_Name, strData) 

	    Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.Static(strObjectName)
		StepStartTime = Time
		If fnc_wait(ActualObject) = "True" Then
				Select Case strAction_Name
					  Case "CHECKEXIST"
							  strExp = strData
							  strAct = ActualObject.Exist(0)
							  If UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The validation message exists as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The validation message does not exist</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  ElseIf UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The validation message does not exist as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The validation message  exists</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
                 End Select
		Else
				UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">" & strObjectName & "does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If
End Function

'==================================================================================================================================================
' Name of the Function     			  : fn_WaitforFewSeconds
' Description       		   		  : This function is used  to give additional wait
' Date and / or Version       	      : 03/24/2015
' Created By              			  : AKshatha
' Example Call						  : Call fn_WaitforFewSeconds("Test Scenario Name","URL") 
'==================================================================================================================================================
Function fn_WaitforFewSeconds(strScenarioName,strData)

         StepStartTime = Time
         strObjectHierarchy = Datatable.value("APP_SCREEN_NAME",strScenarioName)                              
         strObject = Datatable.Value("OBJECT", strScenarioName)
         Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
         
         Wait(strData)
         
 End Function
 
'==================================================================================================================================================
' Name of the Function     			  : OperateOnPbComboBox
' Description       		   		 : This function is used to select the value from combobox
' Date and / or Version       	    : 03/27/2015
' Example Call					 : OperateOnPbComboBox("PbWindow("abc").PbButton(""efg"")","OK", "CLICK", "")
'==================================================================================================================================================
Function OperateOnPbComboBox(strObjectHierarchy, strObjectName, strAction_Name, strData) 

	    Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.PbComboBox(strObjectName)
		StepStartTime = Time
		If fnc_wait(ActualObject) = "True" Then
				Select Case strAction_Name
					  Case "SELECT"
							Wait(1)
							ActualObject.Select strData
							UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"The Value <font color=""blue""><b><i>" & strData &"</i></b></font> is selected from the Combobox <b>"& Mid(strObjectName,4,Len(strObjectName)-3) & "</b> successfully", "Done"
				End Select
		Else
				UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Combobox - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If
End Function

'==================================================================================================================================================
' Name of the Function     			  : OperateOnPbRadioButton
' Description       		   		 : This function is used to enter or retrieve the values on / from the Power Bulider Button object
' Date and / or Version       	    : 19/12/2014
' Example Call					 : OperateOnPbRadioButton("PbWindow("abc").PbButton(""efg"")","OK", "CLICK", "")
'==================================================================================================================================================
Function OperateOnPbRadioButton(strObjectHierarchy, strObjectName, strAction_Name, strData) 

	    Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.PbRadioButton(strObjectName)
		StepStartTime = Time
		
				Select Case strAction_Name
					  Case "SELECT"
							Wait(1)
							ActualObject.Set
							UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "RadioButton <b>"& Mid(strObjectName,4,Len(strObjectName)-3) & "</b> is selected successfully", "Done"
					 
					 Case "CHECKEXIST"
						   strExp = strData
						   strAct = ObjectHierarchy.PbRadioButton(strObjectName).Exist
						   If UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
								UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object exists as expected</font>", "Pass"
						   ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
								UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object does not exist</font>", "Fail"
								Environment.Value("TestStepLog") = "False"
								Environment.Value("TestObjectFlag") = "False"
						   ElseIf UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
								UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object does not exist as expected</font>", "Pass"
						   ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
								UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object exists</font>", "Fail"
								Environment.Value("TestStepLog") = "False"
								Environment.Value("TestObjectFlag") = "False"
						   End If
				End Select
		
End Function


'==================================================================================================================================================
' Name of the Function     			  : OperateOnPbButton
' Description       		   		 : This function is used to enter or retrieve the values on / from the Power Bulider Button object
' Date and / or Version       	    : 19/12/2014
' Example Call					 : OperateOnPbButton("PbWindow("abc").PbButton(""efg"")","OK", "CLICK", "")
'==================================================================================================================================================
Function OperateOnPbButton(strObjectHierarchy, strObjectName, strAction_Name, strData)
   		
		Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.PbButton(strObjectName)
		StepStartTime = Time
		
		If fnc_wait(ActualObject) = "True" Then
					Select Case strAction_Name
					  Case "CLICK"
							  On Error Resume Next
							  Wait(1)
							  ObjectHierarchy.PbButton(strObjectName).Click
							  If Err.Number = 0  Then
							  	 UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Button: <b>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</b> is clicked successfully", "Done"
							  
							  Else	
								 UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Click operation not performed on Button - <b><i>" & strObjectName & "</i></b></font>", "Fail"							  
							  	 Environment.Value("TestStepLog") = "False"
							  End If
					
					Case "CLICKANDSELECT"
							  On Error Resume Next
							  Wait(1)
							  ObjectHierarchy.PbButton(strObjectName).Click
							  If Err.Number = 0  Then
							  	 UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Button: <b>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</b> is clicked successfully", "Done"
							  
							  Else	
								 UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Click operation not performed on Button - <b><i>" & strObjectName & "</i></b></font>", "Fail"							  
							  	 Environment.Value("TestStepLog") = "False"
							  End If
					
					
					 Case "CHECKEXIST"
							  strExp = strData
							  strAct = ActualObject.Exist(0)
							  If UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object exists as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object does not exist</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  ElseIf UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object does not exist as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object exists</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
						Case "CHECKENABLED"
							  blnObjDisable= ActualObject.GetROProperty("enabled")
							 									
								If blnObjDisable = "1" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Disabled as expected</font>", "Pass"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Enabled as expected</font>", "Pass"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Enabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								ElseIf blnObjDisable = "1" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Disabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								End If
						End Select
	           Else
							UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Button - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
							Environment.Value("TestObjectFlag") = "False"
							Environment.Value("TestObjectFlag") = "False"
			End If		
End Function

'==================================================================================================================================================
' Name of the Function     			  : OperateOnPbCheckBox
' Description       		   		 	     : This function is used to select or retrieve the values from the PbCheckBox object
' Date and / or Version       	    : 
' Example Call							 : OperateOnPbCheckBox("Window(""TIPS"").PbCheckBox(""TitleGrid"")","SELECT", "CHECK", "")
'==================================================================================================================================================
Function OperateOnPbCheckBox(strObjectHierarchy, strObjectName, strAction_Name, strData) 

	    Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.PbCheckBox(strObjectName)
		StepStartTime = Time
		If fnc_wait(ActualObject) = "True" Then
				Select Case strAction_Name
						  Case "CHECK"
								On Error Resume Next
								If  ActualObject.GetROProperty("checked") = "OFF" Then
'										ActualObject.Set "ON"
										Wait(1)
										ActualObject.Click
										UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Check box <b>" & Mid(strObjectName,4,Len(strObjectName)-3)&"</b> is checked successfully", "Done"
								Else
										UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Check box <b>" & Mid(strObjectName,4,Len(strObjectName)-3)&"</b> is checked successfully", "Done"
								End If
						  Case "UNCHECK"
								If  ActualObject.GetROProperty("checked") = "ON" Then
'										ActualObject.Set "OFF"
										ActualObject.Click
										UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Check box <b>" & Mid(strObjectName,4,Len(strObjectName)-3)&"</b> is unchecked successfully", "Done"
								 Else
										UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Check box <b>" & Mid(strObjectName,4,Len(strObjectName)-3)&"</b> is unchecked successfully", "Done"
								End If
'***************************** Added by Suresh on 28/07/2011 *************************************************
						Case "CHECKEXIST"
							  strExp = strData
							  strAct = ActualObject.Exist
							  If UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object exists as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object does not exist</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  ElseIf UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object does not exist as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object exists</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
						Case "CHECKENABLED"
								blnObjDisable= ActualObject.GetROProperty("enabled")
                            	If blnObjDisable = "1" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Disabled as expected</font>", "Pass"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Enabled as expected</font>", "Pass"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Enabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								ElseIf blnObjDisable = "1" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Disabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								End If
								Case "CHECKCHECKED"
							 			blnObjDisable= ActualObject.GetROProperty("checked")
                            	If blnObjDisable = "1" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Checked as expected</font>", "Pass"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Not Checked as expected</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Not Checked as expected</font>", "Pass"
                               	ElseIf blnObjDisable = "1" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Checked</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								End If
'**************************************************************************************************************************************
				End Select
		Else
				UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Checkbox - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If
End Function

'==================================================================================================================================================
' Name of the Function     			  : OperateOnPbDataWindow
' Description       		   		  : This function is used to select or retrieve the values from the Power builder DataWindow  object
' Date and / or Version       	   	 : 07/12/2014
' Example Call						: OperateOnPbDataWindow(strObjectHierarchy, strObjectName, strAction_Name, strData)
'==================================================================================================================================================
 Function OperateOnPbDataWindow(strObjectHierarchy, strObjectName, strAction_Name, strData)
	
		'On Error Resume Next
		Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.PBDataWindow(strObjectName)
		StepStartTime = Time
        If fnc_wait(ActualObject) = "True" Then
				Select Case strAction_Name
					 
					 Case "SETCELLDATA"  'This Action is Used to Select the value from List drop down box inside Pb Data Window object
	  						On Error Resume Next
	  						Wait(1)	  						
	  						arrDataVal = Split(strData,";")
	  						If  Instr(strDataRowVal, "VAR_") = 1 Then				' If the value to be taken from already saved variable
								strDataRowVal = Environment.Value(arrDataVal(0))
							Else
	  							strDataRowVal = arrDataVal(0)
	  						End If
	  						strobjName = arrDataVal(1)
	  						strDataVal = arrDataVal(2)
	  						intItemNum = arrDataVal(3)
							
							If  Instr(strDataRowVal, "VAR_") = 1 Then				' If the value to be taken from already saved variable
								strDataRowVal = Environment.Value(arrDataVal(0))
							End If
							
	  						ActualObject.SelectCell  strDataRowVal,strobjName 
							ActualObject.SetCellData strDataRowVal, strobjName,strDataVal						
							ActualDataVal = ActualObject.GetCellData(strDataRowVal, strobjName)    
							
							If intItemNum = ActualDataVal  Then
									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value: " & "<font color=""blue""><b><i>" & strDataVal & "</i></b></font> is selected from the Dropdown <b>"  & Mid(strObjectName,4,Len(strObjectName)-3) & "</b>", "Done"
							ElseIf intItemNum <> ActualDataVal Then
  									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>","<font color=""red"">" & "The value:  <b><i>" & strDataVal & "</i></b> is not available in the Dropdown <b>"  & Mid(strObjectName,4,Len(strObjectName)-3) &"</b></font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							ElseIf Err.Number <> 0 Then
							  		UpdateReport "TESTSTEP", "","<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>" , "<font color=""red"">"&  Time & "</font>", "<font color=""red"">The Value: <b><i>" & strDataVal & "</i></b> is not selected from List Object<b>" & Mid(strObjectName,4,Len(strObjectName)-3) &"</b> Due to Error" & Err.Description & "</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							End If	
					
					Case "LISTVALUESELECT"  'This Action is Used to Select the value from List drop down box inside Pb Data Window object if object is already focude/selected this case is similar to SETCELLDATA without seelctcell method
	  						On Error Resume Next
	  						Wait(1)	  						
	  						arrDataVal = Split(strData,";")
	  						If  Instr(strDataRowVal, "VAR_") = 1 Then				' If the value to be taken from already saved variable
								strDataRowVal = Environment.Value(arrDataVal(0))
							Else
	  							strDataRowVal = arrDataVal(0)
	  						End If
	  						strobjName = arrDataVal(1)
	  						strDataVal = arrDataVal(2)
							
							If  Instr(strDataRowVal, "VAR_") = 1 Then				' If the value to be taken from already saved variable
								strDataRowVal = Environment.Value(arrDataVal(0))
							End If
							
							ActualObject.SetCellData strDataRowVal, strobjName,strDataVal						
							
							If Err.Number = 0   Then
									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value: " & "<font color=""blue""><b><i>" & strDataVal & "</i></b></font> is selected from the Dropdown <b>"  & Mid(strObjectName,4,Len(strObjectName)-3) & "</b>", "Done"
							ElseIf Err.Number <> 0 Then
							  		UpdateReport "TESTSTEP", "","<font color=""red"">" & Environment.Value("strDescription") & "</font>","<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>","<font color=""red"">" & "The Value: <b><i>" & strDataVal & "</i></b> is not selected from List Object<b>" & Mid(strObjectName,4,Len(strObjectName)-3) &"</b> Due to Error" & Err.Description & "</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							Else  
  									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>","<font color=""red"">" & "The value:  <b><i>" & strDataVal & "</i></b> is not available in the Dropdown <b>"  & Mid(strObjectName,4,Len(strObjectName)-3) & "</b></font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							End If
					
                    Case "SETRADIOBUTTON" 'This Action is used to set the radio button ON inside Pb Data Window object
	  						On Error Resume Next
	  						Wait(1)	  						
	  						arrDataVal = Split(strData,";")
	  						If  Instr(strDataRowVal, "VAR_") = 1 Then				' If the value to be taken from already saved variable
								strDataRowVal = Environment.Value(arrDataVal(0))
							Else
	  							strDataRowVal = arrDataVal(0)
	  						End If
	  						strobjName = arrDataVal(1)
	  						strDataVal = arrDataVal(2)

							ActualObject.SetCellData strDataRowVal, strobjName,strDataVal						
							ActualDataVal = ActualObject.GetCellData(strDataRowVal, strobjName)    
							
							If strDataVal = ActualDataVal Then
   								UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "RadioButton <b><i>"& Mid(strObjectName,4,Len(strObjectName)-3) & "</i></b> is selected successfully", "Done"
					  		
					  		Else
								UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The RadioButton - <b><i>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</i></b> does not exist" & "</font>", "Fail"
								Environment.Value("TestStepLog") = "False"
							End If	
							
					Case "SELECTCELL"	'This Action is used to Select the object inside Pb data Window object(Focus on the object inside Pb Data window object or click on the object inside Pb Data window object) to perform any action on it 
	  						On Error Resume Next
	  						Wait(1)	  						
	  						arrDataVal = Split(strData,";")
	  						strDataRowVal = arrDataVal(0)
	  						If  Instr(strDataRowVal, "VAR_") = 1 Then				' If the value to be taken from already saved variable
								strDataRowVal = Environment.Value(arrDataVal(0))
							Else
	  							strDataRowVal = arrDataVal(0)
	  						End If
	  						strobjName = arrDataVal(1)
	  						
	  						If UBound(arrDataVal) > 2 Then
	  							strDataVal = arrDataVal(2)
								ObjectHierarchy.Activate  						
		  						ActualObject.SelectCell  strDataRowVal,strobjName 													
								strActualDataVal = ActualObject.GetCellData(strDataRowVal, strobjName)    
								
								If (strDataVal = strActualDataVal) AND (Err.Number = 0) Then
								  	UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value: <font color=""blue""><b><i>" & strDataVal & "</i></b></font> is selected from <b>"  & Mid(strObjectName,4,Len(strObjectName)-3) & "</b>" & " in PbDataWindow", "Done"
								Else
  									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>","<font color=""red"">" & "The value:  <b><i>" & trim(strDataVal) & "</i></b> is not available in the List Object <b>" & strobjName & "</b> Due to Error "& Err.Description & "</font>" , "Fail"
									Environment.Value("TestStepLog") = "False"
							    End If
							
							Else
								ObjectHierarchy.Activate  						
		  						ActualObject.SelectCell  strDataRowVal,strobjName 							
					  			UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Object <b><i>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</i></b> is selected/Focused in the PbDataWindow" , "Done"
				  			End IF
				  			
					  Case "TYPEVALUE" 'This Actionis used to Type Value into Edit field inside Pb Data window object
							On Error Resume Next  
						  	Wait(1)  							  
							arrDataVal = Split(strData,";")
							
							If Instr( arrDataVal(0), "VAR_") = 1 Then
								strDataRowVal = Environment.Value(arrDataVal(0))
							Else
								strDataRowVal = arrDataVal(0)
							End If 
							 strobjName = arrDataVal(1)
							   
							  If  Instr(arrDataVal(2), "VAR_") = 1 Then				'If the value to be taken from already saved variable
								  strDataVal = Environment.Value(arrDataVal(2))
							  Else
							  	  strDataVal = arrDataVal(2)
							  End If
							   
							  ActualObject.SelectCell  strDataRowVal,strobjName 
							  ActualObject.Type strDataVal							  
							  strActualDataVal = ActualObject.GetCellData(strDataRowVal, strobjName)
							 
							If Instr(1,strActualDataVal, "/") <> 0 AND Instr(1,strDataVal, "/") <> 0  Then 'To compare Date Format from Application and Expected
							 	
							 	If StrComp(CDate(strActualDataVal),CDate(strDataVal)) = 0 Then
									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Value: <font color=""blue""><b><i>" & strDataVal & "</i></b></font> is entered in field <b>" & Mid(strObjectName,4,Len(strObjectName)-3) &"</b> successfully", "Done"							  					  
							   	End If
							  	
							Else If StrComp(Trim(strActualDataVal),Trim(strDataVal)) = 0 Then 'To Compare String Data from Actual And Expected
									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Value: <font color=""blue""><b><i>" & strDataVal & "</i></b></font> is entered in field <b>" & Mid(strObjectName,4,Len(strObjectName)-3) &"</b> successfully", "Done"
							    
							Else
									UpdateReport "TESTSTEP", "","<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">" & "The Value: <b><i>" & strDataVal & "</i></b> is not entered in field <b>" & Mid(strObjectName,4,Len(strObjectName)-3) &"</b> Due to Error " & Err.Description & "</font>", "Fail"
							  		Environment.Value("TestStepLog") = "False"
							End If
							End If
							
					Case "VALIDATEDATAEXISTANCEINGRID" 'This Case is used to Check Data Exist in Pb Data window Grid object
						  
						  On Error Resume Next   
						  RowCount = ActualObject.RowCount  
        	 			  strDataVal = strData
        	 			  If RowCount = 0 Then         'When data is not present in datawindow
                			 
                			 If strDataVal = "TRUE" Then   'When Testcase expects data to be  present
			                   UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Data is not populated in <b>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</b> datawindow Grid</font>", "Fail"
			                   Environment.Value("TestStepLog") = "False"
			                  
               			 	 ElseIf strDataVal = "FALSE" Then             'When Testcase expects no rows to be retrieved
                   				UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Data is not populated in <b>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</b> datawindow Grid</font>", "Pass"
               			 	 End If
                
            			 ElseIf RowCount > 0 Then                             'When data is present in datawindow
                
			                 If strDataVal = "TRUE" Then                     'When Testcase expects data to be present
			                 	UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Data is populated in <b>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</b> datawindow Grid</font>", "Pass"
			                 ElseIf strDataVal = "FALSE" Then               'When Testcase expects no rows to be retrieved
			                    UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Data is populated in <b>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</b> datawindow Grid</font>", "Fail"
			                    Environment.Value("TestStepLog") = "False"
			                 End If
                 
             			End If
							  
					Case "SETVALUE" 'This Case is used to SetValue into Edit field inside Pb Data window object without using SelectCell method on Object
							
							On Error Resume Next   
						  	Wait(1)  							  
							arrDataVal = Split(strData,";")
							If Instr( arrDataVal(0), "VAR_") = 1 Then
								strDataRowVal = Environment.Value(arrDataVal(0))
							Else
								strDataRowVal = arrDataVal(0)
							End If
							 strobjName = arrDataVal(1)
							 strDataVal = arrDataVal(2)
							  
							  If  Instr(arrDataVal(0), "VAR_") = 1 Then				' If the value to be taken from already saved variable
								  strDataRowVal = Environment.Value(arrDataVal(0))
								  strDataRowVal = "#" & strDataRowVal
							  End If
							  
							  ActualObject.Type strDataVal							  
							  strActualDataVal = ActualObject.GetCellData(strDataRowVal, strobjName)
							 
							If Instr(1,strActualDataVal, "/") <> 0 AND Instr(1,strDataVal, "/") <> 0  Then 'To compare Date Format from Application and Expected
							 	
							 	If StrComp(CDate(strActualDataVal),CDate(strDataVal)) = 0 Then
									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Value: <font color=""blue""><b><i>" & strDataVal & "</i></b></font> is entered in field <b>" & Mid(strObjectName,4,Len(strObjectName)-3) &"</b> successfully", "Done"							  					  
							   	End If
							  	
							Else If StrComp(Trim(strActualDataVal),Trim(strDataVal)) = 0 Then 'To Compare String Data from Actual And Expected
									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Value: <font color=""blue""><b><i>" & strDataVal & "</i></b></font> is entered in field <b>" & Mid(strObjectName,4,Len(strObjectName)-3) &"</b> successfully", "Done"
							    
							Else
									UpdateReport "TESTSTEP", "","<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>","<font color=""blue"">" & "The Value: <b><i>" & strDataVal & "</i></b> is not entered in field <b>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</b></font>", "Fail"
							  		Environment.Value("TestStepLog") = "False"
							End If
							End If	
							
					Case "CLEAREFIELDANDTYPE" 'This Case is used to SetValue new data by clearing old Data into Edit field  inside Pb Data window object without using SelectCell method on Object
						  
						  On Error Resume Next 
						  Wait(1)	  
	  					  arrDataVal = Split(strData,";")
	  					  
	  					  If Instr( arrDataVal(0), "VAR_") = 1 Then
								strDataRowVal = Environment.Value(arrDataVal(0))
						  Else
								strDataRowVal = arrDataVal(0)
						  End If
	  					  strobjName = arrDataVal(1)
	  					  
	  					  If Instr( arrDataVal(2), "VAR_") = 1 Then
								strDataVal = Environment.Value(arrDataVal(0))
						  Else
								strDataVal = arrDataVal(2)
						  End If
	  					 
	  					 ActualObject.SelectCell  strDataRowVal,strobjName 
						  Wait(1)							  
	  					  Set WshShell = CreateObject("WScript.Shell")    
		                      WshShell.SendKeys "^+{DELETE}"
		                      WshShell.SendKeys "+"&"{END}"
		                      WshShell.SendKeys "{DELETE}"
		                      wait(2)
						  ActualObject.Type strDataVal							  
						  strActualDataVal = ActualObject.GetCellData(strDataRowVal, strobjName)
						 
						If Instr(1,strActualDataVal, "/") <> 0 AND Instr(1,strDataVal, "/") <> 0  Then 'To compare if Actual data is Date type  from Application and Expected
							 	
							 	If StrComp(CDate(strActualDataVal),CDate(strDataVal)) = 0 Then
									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Value: <font color=""blue""><b><i>" & strDataVal & "</i></b></font> is entered in field <b>" & Mid(strObjectName,4,Len(strObjectName)-3) &"</b> successfully", "Done"							  					  
							   	End If
							  	
							Else If StrComp(Trim(strActualDataVal),Trim(strDataVal)) = 0 Then 'To Compare if Data is String Data type with Application Data And Expected Data
									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Value: <font color=""blue""><b><i>" & strDataVal & "</i></b></font> is entered in field <b>" & Mid(strObjectName,4,Len(strObjectName)-3) &"</b> successfully", "Done"
							    
							Else
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">" & "The Value: <b><i>" & strDataVal & "</i></b> is not entered in field <b>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</b></font>", "Fail"
							  		Environment.Value("TestStepLog") = "False"
							End If
							End If
							
					Case "CUSTOMEFIELDTYPE" 'This Case is used to SetValue new data into Edit field by clearing old Data in field inside Pb Data window object without using SelectCell method on Object
						 
						 On Error Resume Next  
	  					  arrDataVal = Split(strData,";")
	  					  
	  					  If  Instr(arrDataVal(0), "VAR_") = 1 Then				' If the value to be taken from already saved variable
								  strDataRowVal = Environment.Value(arrDataVal(0))
						  Else
						  		 strDataRowVal = arrDataVal(0)
						  End If
						  strobjName = arrDataVal(1)
						  
	  					  If  Instr(arrDataVal(2), "VAR_") = 1 Then				' If the value to be taken from already saved variable
								  strDataVal = Environment.Value(arrDataVal(2))
						  Else
						  		 strDataVal = arrDataVal(2)
						  End If
	  					  
						  Wait(1)							  
	  					  Set WshShell = CreateObject("WScript.Shell")    
		                      WshShell.SendKeys "+"&"{END}"
		                  Wait(1)    
						  ActualObject.Type strDataVal							  
						  
						  If Err.Number = 0 Then
						  	 UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Value: <font color=""blue""><b><i>" & strDataVal & "</i></b></font> is entered in field <b>" & Mid(strObjectName,4,Len(strObjectName)-3) &"</b> successfully", "Done"
						  Else
							UpdateReport "TESTSTEP", "","<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>","<font color=""red"">" &  "The Value: <b><i>" & strDataVal & "</i></b> is not entered in field <b>" & Mid(strObjectName,4,Len(strObjectName)-3) &"</b> Due to Error <b>" & Err.Description & "</b></font>" , "Fail"
					  		Environment.Value("TestStepLog") = "False"
						  End If  
						  
					Case "GETCELLDATA" 'This case is used to store value into variable for later usage during runtime from application
						  
						  On Error Resume Next
							 arrDataval = Split(strData,";")
							 
							 If  Instr(arrDataVal(0), "VAR_") = 1 Then				' If the value to be taken from already saved variable
								  strDataRowVal = Environment.Value(arrDataVal(0))
						  	Else
						  		 strDataRowVal = arrDataVal(0)
						  	End If
							 strobjName = arrDataval(1)
							 
							 Environment.Value("VAR_"& arrDataval(2)) = ActualObject.GetCellData(strDataRowVal, strobjName) 'store the Data into variable Name with prefix VAR_
 							 If Err.Number = 0 Then							 
								 UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value: <i>"& Environment.Value("VAR_"& arrDataval(2))& "</i> is stored", "Done"
							 Else
							 	 UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>","<font color=""red"">" &  "The Value: <b><i>" & strDataVal & "</i></b> is not Stored from field <b>" & Mid(strObjectName,4,Len(strObjectName)-3) &"</b> Due to Error <b>" & Err.Description & "</b></font>" , "Fail"
						  		 Environment.Value("TestStepLog") = "False"
							 End If
                   
                   Case "GETCELLDATAANDCOMPAREVALUE" 'This Case is used to compare the Cell Value in the Pb Data window grid object  with Expected Data and if need row number can be stored in variable
							 
							 On Error resume Next
							 Wait(2)
							 arrDataval = Split(strData,";")
							 strColumnName = arrDataval(0)
							 
							 If  Instr(arrDataval(1), "VAR_") = 1 Then				' If the value to be taken from already saved variable
								  strExpVal = Environment.Value(arrDataval(1))
						  	Else
						  		 strExpVal = arrDataval(1)
						  	End If
							strFirstColumn = arrDataval(2)
							 
							 If UBound(arrDataval) > 2 Then ' Passing the variable name TO Store the ROw Number to use later  
							 	Environment.Value("VAR_"& arrDataval(3)) = arrDataval(3)
							 ElseIf Instr(1,arrDataval(1),"VAR") = 1 Then ' TO use already stored value as expected value for comparision
							 		strExpVal = Environment.Value(arrDataval(1))
							 End If
							 
							 RowCount = ActualObject.RowCount						 
						 	 
						 	For intRowCount = 1 To RowCount Step 1						 	 
						 	 	ActualObject.SelectCell "#"&intRowCount,strFirstColumn
						 		strActualDataVal = ActualObject.GetCellData("#"&intRowCount,strColumnName)  
						 		If Instr(1,strActualDataVal, "/") <> 0 AND Instr(1,strExpVal, "/") <> 0  Then
							 	
								 	If StrComp(CDate(strActualDataVal),CDate(strExpVal)) = 0 Then 'To compare Date Data Type Data if the Actual Data is Date format. convert actual and expected data to Date format and comapre with each other
										blnflag = True
								   	End If
								ElseIf StrComp(Trim(strActualDataVal),Trim(strExpVal)) = 0 Then 'To compare string Data Type Data if the Actual Data is String type. convert actual and expected data to string format and comapre with each other
							    	blnflag = True
								End If
								
							 	If blnflag Then
							 	  	blnflag = True
							 	    Exit For						 		
							 	End If
						 	Next
								Environment.Value("VAR_"& arrDataval(3)) = intRowCount 
							 If blnflag  Then
							 		UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Expected Value - "& strExpVal &" and Actual value - "& strExpVal & " are matching</font>", "Pass"
							 Else
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Value mismatch, Expected Value is: <i>" & strExpVal & "</i>, and Actual value on the application is: <i>" & strActual & "</i></font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							 End If							 								 
					
					Case "STOREVALUE" 'This Case is used to Store the value fetched form any object into Environmental variable it takes variable as parameter where it stores value
							 
							 On Error Resume Next
							 arrDataval = Split(strData,";")
							
							 If  Instr(arrDataVal(0), "VAR_") = 1 Then				' If the value to be taken from already saved variable
								  strDataRowVal = Environment.Value(arrDataVal(0))
						  	 Else
						  		 strDataRowVal = arrDataVal(0)
						  	 End If
							 strobjName = arrDataval(1)
							 Environment.Value("VAR_"& arrDataval(2)) = arrDataval(2)
							 Environment.Value("VAR_"& arrDataval(2)) =ActualObject.GetCellData(strDataRowVal, strobjName)
							
							If Err.Number = 0 OR  (Environment.Value("VAR_"& arrDataval(2)) <> "") Then
								 UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value: <i>"& Environment.Value("VAR_"& arrDataval(2))& "</i> is stored in the variable <b><i>" & arrDataval(2) &"</i></b>", "Done"
					  		 Else
					  		 	 UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Value is not stored into Environmental Variable <i>" & arrDataval(2) & "</i></font>", "Fail"
									Environment.Value("TestStepLog") = "False"
					  		 End If		
					  		  
					Case "ACTIVATECELLINGRID" 'This Case is used to Activate the Window from the selected list Value/record in Pb Data window Grid to update the data 
							 
							 On Error Resume Next
							 Wait(2)
							 arrDataval = Split(strData,";") 
							 strColumnName = arrDataval(0)
							 
							 If  Instr(arrDataVal(1), "VAR_") = 1 Then				' If the value to be taken from already saved variable
								  strExpVal = Environment.Value(arrDataVal(1))
						  	 Else
						  		 strExpVal = arrDataVal(1)
						  	 End If
							 strFirstColumn = arrDataval(2)
							 
							 RowCount = ActualObject.RowCount						 
						 	 
						 	For intRowCount = 1 To RowCount Step 1						 	 
						 	 	
						 	If ActualObject.GetCellData("#"&intRowCount,strColumnName) = strExpVal Then
						 		ActualObject.ActivateCell "#"&intRowCount,strFirstColumn
						 	  	blnflag = True
						 	    Exit For						 		
						 	End If
						 	Next
							 
							 If blnflag Then
							 		UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Expected Cell Data - "& strExpVal &" and Actual Cell Data in Grid - "& strActual & " are matching, Cell is selected and activated</font>", "Pass"
							 Else
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Value mismatch, Expected Cell Data: <i>" & strExpVal & "</i>, and Actual and Actual Cell Data in Grid: <i>" & strActual & "</i></font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							 End If	 
					 
					 
					 Case "COMPAREVALUE"		'This Case is used to Compare the Actual Data And Expected Data form the object inside PB Data Window					 
							On Error Resume Next
							arrDataVal = Split(strData,";")
							
							If Instr( arrDataVal(0),"VAR") <> 0Then
								strDataRowVal =	Environment.Value(arrDataval(0))
								
							Else
								strDataRowVal = arrDataVal(0)							
							End If
								
							strobjName = arrDataVal(1)
							strExpVal = arrDataVal(2)
						     
						    If Instr(1,arrDataval(2),"VAR") = 1 Then ' TO use already stored value as expected value for comparision
							 		strExpVal = Environment.Value(arrDataval(2))		
						    ElseIf UBound(arrDataVal) = 3 Then
								strExpVal = arrDataVal(3)
						    
						    End If						 
							  
						  	If ActualObject.GetCellData("#"&strDataRowVal, strobjName) = ""  Then		' Checking if the first List Value/record in the Pb Data window Grid is Empty						  
							    strActual = ActualObject.GetCellData("#2", strobjName)							  
						 	Else								  
							     strActual = ActualObject.GetCellData(strDataRowVal, strobjName)	
						 	End If
							  
							If Instr(1, strActual, ".0") <> 0 Then 'checking if the Data containg double data type value from actual Data 
								   strActual = cint(strActual)
							ElseIf Instr(1, strActual,"/") <> 0  Then			
								   strExpVal = CDate(strExpVal)		
								   strActual = CDate(strActual)
							ElseIf Instr(1, strActual,"/") <> 0 AND len(strActual) > 10 Then 'checking Data is Date data Type format if Data format convert actual and expected into date format
								  strActual = CDate(Left(strActual,10))
								  strExpVal = CDate(strExpVal)								 							
							End If
							 
							If (Trim(UCASE(strExpVal)) = Trim(UCASE(strActual)) OR (strExpVal = strActual)) AND Err.Number = 0 Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Expected Value - "& strExpVal &" and Actual value - "& strActual & " are matching</font>", "Pass"
							 Else
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Value mismatch, Expected Value is: <i>" & strExpVal & "</i>, and Actual value on the application is: <i>" & strActual & "</i></font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							 End If
					
					Case "CLICK" 'This Case is used to click on the Object inside Pb data window 
					 		On Error Resume Next				
	  					    arrDataVal = Split(strData,";")
	  		                If Instr( arrDataVal(0), "VAR_") = 1 Then
								strDataRowVal = Environment.Value(arrDataVal(0))
							Else
								strDataRowVal = arrDataVal(0)
							End If
	  				  		strDataColVal = arrDataVal(1)
	
	  						ActualObject.SelectCell  strDataRowVal,strDataColVal	
							
							If Err.Number = 0 Then
								  	UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Object <font color=""blue""><b><i>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</i></b></font> is Clicked", "Done"
							Else
  									UpdateReport "TESTSTEP", "","<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Object <b><i>" & Mid(strObjectName,4,Len(strObjectName)-3) & " </i></b> is not Clicked successfully </font>" , "Fail"
									Environment.Value("TestStepLog") = "False"
							End If
							
					Case "SELECTLISTVALUE" 'This case is used to Seelct the List Value/Record in the PB Data Window grid Object 
						 
						 On Error Resume Next
						 arrDataVal = Split(strData,";")
						 strColumnName = arrDataVal(0)
						 
						 If Instr( arrDataVal(1), "VAR_") = 1 Then
								strExpVal = Environment.Value(arrDataVal(1))
						 Else
								strExpVal = arrDataVal(1)
						 End If
						 
						 RowCount = ActualObject.RowCount 'get the Row Count 
						 
						 For intRowCount = 1 To RowCount Step 1
						 	 
						 	ActualObject.SelectCell "#"&intRowCount,"#1" 
						 	
						 	If ActualObject.GetCellData("#"&intRowCount,strColumnName) = strExpVal Then 'Validate if Acutal Cell Data in the Grid is Equals expected Data
						 	   blnflag = True
						 	   Exit For						 		
						 	End If
						 Next
						 
						 If blnflag AND Err.Number = 0 Then
						 	UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value: <font color=""blue""><b><i>" & strExpVal & "</i></b></font> is selected from the List <b>"  &Mid(strObjectName,4,Len(strObjectName)-3)& "</b>", "Done"
						 Else
						 	UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The value: <b><i>" & strExpVal & "</i></b> is not available in the List <b>"  &Mid(strObjectName,4,Len(strObjectName)-3) &"</b></font>", "Fail"
							Environment.Value("TestStepLog") = "False"
						 End If
					
					Case "COMPARESPECVALUEORRANGEINCOL"  'This case is used to 
							
							On Error Resume Next 
							 Wait(2)
							 arrDataval = Split(strData,";")
							 strColumnName = arrDataval(0)
							 strExpVal = arrDataval(1)
							 strFirstColumn = arrDataval(2)
							 
							 RowCount = ActualObject.RowCount	
	
						 	For intRowCount = 1 To RowCount Step 1						 	 
						 	 	ActualObject.SelectCell "#"&intRowCount,strFirstColumn
						 		strActualDataVal = ActualObject.GetCellData("#"&intRowCount,strColumnName)  
						 		blnflag = false
						 		rangeflag = true
						 		If UBOund(arrDataVal) > 2 Then
	  									strDataVal = arrDataVal(3)
	  									rangeflag = false
						 		 	If Instr(1,strActualDataVal, "/") <> 0 AND Instr(1,strExpVal, "/") <> 0  Then
							 	
								 		strActualDataVal = CDate(strActualDataVal)
								 		strExpVal = CDate(strExpVal) 
									End If		
						 		 	If ((strExpVal <= strActualDataVal) AND (strDataVal >= strActualDataVal)) Then
						 		 	 	blnflag = true
						 		 	End If
								 ElseIf StrComp(Trim(strActualDataVal),Trim(strExpVal)) = 0 Then 							    		
								 		blnflag = true
								 End If
						 	Next
							 
							 If blnflag AND Err.Number = 0 Then
							    If rangeflag Then
							 		UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Expected Value - "& strExpVal &" and Actual value - "& strActualDataVal & " are matching</font>", "Pass"
							 	Else
							        UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green""> Data is within the range "& strExpVal &" and "& strDataVal &"</font>", "Pass"
							    End If
							 Else
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Value mismatch, Expected Value is: <i>" & strExpVal & "</i>, and Actual value on the application is: <i>" & strActualDataVal & "</i></font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							 End If
			 						
					Case "COMPAREVALUEFORSPECIFICROW" 'This Case is used to compare the specific column data value for specific row data
						 
						 On Error Resume Next 
						 arrDataVal = Split(strData,";")
						 strSpecColumnName = arrDataVal(0)
						 strSpecRowVal = arrDataVal(1)
						 strColumnName = arrDataval(2)
						 
						 If Instr( arrDataVal(3), "VAR_") = 1 Then
								strExpVal = Environment.Value(arrDataVal(3))
						 Else
								strExpVal = arrDataVal(3)
						 End If
						 
						 RowCount = ActualObject.RowCount 'Get thwe row count in PB Data Grid
						 
						 For intRowCount = 1 To RowCount Step 1
						 	 
						 	ActualObject.SelectCell "#"&intRowCount,strSpecColumnName
						 	blnflag = false
							 	If ActualObject.GetCellData("#"&intRowCount,strSpecColumnName) = strSpecRowVal Then 'Compare unique Column Data and proceed to next columns in the same row 
							 	   strActualDataVal = ActualObject.GetCellData("#"&intRowCount,strColumnName)  
							 		If Instr(1,strActualDataVal, "/") <> 0 AND Instr(1,strExpVal, "/") <> 0  Then 'If the Actual Value is Date format and expected from data datasheet is date which treated as string is converted into CDate format
								 	
									 	If StrComp(CDate(strActualDataVal),CDate(strExpVal)) = 0 Then
											blnflag = True
											Exit For
									   	End If
								  
									ElseIF StrComp(Trim(strActualDataVal),Trim(strExpVal)) = 0 Then 'If the Actual Value and the expceted values are string 
							    		blnflag = True
							    		Exit For
									End If    	
								 End If
						 	Next
							 
							 If blnflag AND Err.Number = 0 Then
							 		UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Expected Value - "& strExpVal &" and Actual value - "& strActualDataVal & " are matching</font>", "Pass"
							 Else
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Value mismatch, Expected Value is: <i>" & strExpVal & "</i>, and Actual value on the application is: <i>" & strActual & "</i></font>", "Fail"
									Environment.Value("TestStepLog") = "False"
									Environment.Value("TestObjectFlag") 
							 End If	
			 	
			 	End Select
        Else
                UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The PbDataWindow Object - <b><i>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If
		
End Function


'==================================================================================================================================================
' Name of the Function     			  : CreateHtmlReport
' Description       		   		 	     : This function should be called in the beginning of the execution in order to create the result file
' Date and / or Version       	    : 
' Example Call							 : CreateHtmlReport()
'==================================================================================================================================================
Public Function CreateHtmlReport()
	Dim fso, ofile 
	DateTime = Date &"_" &Time
	set regEx = New RegExp
	regEx.global = true
	regEx.pattern = "[/\ \:]"
	TimeStamp = regEx.replace(DateTime, "")
	strIpAddr = Environment.Value("LocalHostName")
	strUserName = Environment.Value("UserName")
	
	Environment.Value("HTMLFoldername") = Environment.Value("strTestCase") & "_" & TimeStamp
	Environment.Value("HTMLPath") = Environment.Value("ResultPath") & Environment.Value("HTMLFoldername")

	Set objHTML = CreateObject("Scripting.FileSystemObject")
	If objHTML.FolderExists (Environment.Value("HTMLPath")) Then
		objHTML.DeleteFolder(Environment.Value("HTMLPath"))
		objHTML.CreateFolder(Environment.Value("HTMLPath"))
	Else
		objHTML.CreateFolder(Environment.Value("HTMLPath"))
	End If
	
	objHTML.CreateFolder(Environment.Value("HTMLPath") & "\Snapshot")
    Environment.Value("strReportPath") = Environment.Value("ResultPath") & Environment.Value("HTMLFoldername") & "\TestAutomation_Execution_Report.html"
	
	Set fso = CreateObject("Scripting.FileSystemObject") 'set fso as FileSystemObjectobject ,  create text file and assign it to object ofile
	set ofile = fso.CreateTextFile(Environment.Value("strReportPath"),True) 'Replace if already exists
		ofile.Write "<html>"     'writing content in HTML format 
		ofile.WriteLine "<body style="&"""background-color:#EEEEEE;"""&">"
		ofile.WriteLine "<h1 style="&"""text-align:center"""&">"
        ofile.WriteLine "<p style="&"""color:NAVY"""&">"
		ofile.WriteLine "<u>"
		ofile.WriteLine gstrReportHeader
		ofile.WriteLine "</u>"
		ofile.WriteLine "</p>"
		ofile.WriteLine "</h1>"
		ofile.WriteLine "<table border=""4"" style=""font-family: verdana; font-size: 12px;  font-weight: normal ; text-align:left"" align=""center"">"
		ofile.WriteLine "<caption>Test Execution Summary</caption>"
		ofile.WriteLine "<tr>"
		ofile.WriteLine "<td bgcolor=""#BBBBBB"" WIDTH=""70%"">Total Test Scenarios Executed</td>"
		ofile.WriteLine "<td align=""center""></td>"
		ofile.WriteLine "</tr>"
		ofile.WriteLine "<tr>"
		ofile.WriteLine "<td bgcolor=""#BBBBBB"" WIDTH=""70%"">Total Pass Count</td>"
		ofile.WriteLine "<td align=""center"" Style="&"""Color:green"""&"></td>"
		ofile.WriteLine "</tr>"
		ofile.WriteLine "<tr>"
		ofile.WriteLine "<td bgcolor=""#BBBBBB"" WIDTH=""70%"">Total Fail Count</td>"
		ofile.WriteLine "<td align=""center"" Style="&"""Color:red"""&"></td>"
		ofile.WriteLine "</tr>"
		ofile.WriteLine "<tr>"
		ofile.WriteLine "<td bgcolor=""#BBBBBB"" WIDTH=""70%"">Start Time</td>"
		ofile.WriteLine "<td align=""center"" Style="&"""Color:#264D73"""&">"&Now&"</td>"
		ofile.WriteLine "</tr>"
		ofile.WriteLine "<tr>"
		ofile.WriteLine "<td bgcolor=""#BBBBBB"" WIDTH=""70%"">End Time</td>"
		ofile.WriteLine "<td align=""center"" Style="&"""Color:#264D73"""&"></td>"
		ofile.WriteLine "</tr>"
		ofile.WriteLine "<tr>"
		ofile.WriteLine "<td bgcolor=""#BBBBBB"" WIDTH=""70%""> IP Address </td>"
		ofile.WriteLine "<td align=""center"" Style="&"""Color:#264D73"""&">"& strIpAddr &"</td>"
		ofile.WriteLine "</tr>"
		ofile.WriteLine "<tr>"
		ofile.WriteLine "<td bgcolor=""#BBBBBB"" WIDTH=""70%""> User Name </td>"
		ofile.WriteLine "<td align=""center"" Style="&"""Color:#264D73"""&">" & strUserName & "</td>"
		ofile.WriteLine "</tr>"
		ofile.WriteLine "</table>"
		ofile.WriteLine "<table border="&"""0"""&" style="&"""text-align:center"""&" align="&"""center"""&">"
		ofile.WriteLine "<caption></caption>"
		ofile.WriteLine "<tr></tr>"
		ofile.WriteLine "<tr></tr>"
		ofile.WriteLine "<tr></tr>"
		ofile.WriteLine "<tr></tr>"
		ofile.WriteLine "</table>"
		ofile.WriteLine "<table border="&"""2"""&" style="&"""text-align:left"""&" align="&"""center"""&">"
		ofile.WriteLine "<tr>"
		ofile.WriteLine "<th align="&"""center"""&" bgcolor=""#67799"" WIDTH=""30%"" Style=""Color:white"">Step Description</th>"
		ofile.WriteLine "<th align="&"""center"""&" bgcolor=""#67799"" WIDTH=""40%"" Style=""Color:white"">Step Executed</th>" 
		ofile.WriteLine "<th align="&"""center"""&" bgcolor=""#67799"" WIDTH=""6%"" Style=""Color:white"" >Status</th>"
		ofile.WriteLine "<th align="&"""center"""&" bgcolor=""#67799"" WIDTH=""12%"" Style=""Color:white"" >Start Time</th>" 
		ofile.WriteLine "<th align="&"""center"""&" bgcolor=""#67799"" WIDTH=""12%"" Style=""Color:white"" >End Time</th>" 
		ofile.WriteLine "</tr>"
	ofile.close  'closing file
	
	Set ofile = Nothing 'freeing object memory for ofile and fso
	Set fso= Nothing
	Set regEx = Nothing
End Function

'==================================================================================================================================================
' Name of the Function     			  : UpdateReport
' Description       		   		 	     : This function is used to update the TESTCASE or TESTSTEP result
' Date and / or Version       	    : 
' Example Call							 : UpdateReport "TESTSTEP", "", "Enter the user Name", StartTime, EndTime, "User Name is entered", "Pass"
'==================================================================================================================================================
Public Function UpdateReport(VarStepDetails,VarTestCase,VarDesc,VarStartTime,VarEndTime,VarActRes,VarStatus)
	Dim fso, Ofile, intTestSNo

    Set regEx = New RegExp					' Added by Shrinidhi 8-Jul-11 to replace the special characters in the test case name / description / actual result
	regEx.global = true
	regEx.pattern = "['&  ]"
	VarDesc = regEx.replace(VarDesc, " ")
	VarActRes = regEx.replace(VarActRes, " ")
	VarTestCase = regEx.replace(VarTestCase, " ")

	Set fso = CreateObject("Scripting.FileSystemObject") 	'Creating fso object
	set Ofile = fso.OpenTextFile(Environment.Value("strReportPath"),8) 'Opening file in append mode (8). appends lines at end of file.
		Ofile.writeLine "<tr>" 	'Writing results in html format , write new row in table
		Select Case(VarStepDetails)
			Case "TESTCASE"
				UpdateLog "Executing the Test Case: " & VarTestCase
				Ofile.writeLine "<td colspan = ""100%"" bgcolor=""#E9967A"" style=""font-family: Times New Roman;font-weight: bold;weight:700"">"
				Ofile.writeLine "<TestDescription>" & VarTestCase & "</TestDescription>"
				Ofile.writeLine "</td>"
				
			Case "TESTSTEP"
				UpdateLog "Actual step executed: " &	VarActRes & ", Step result: " & VarStatus
				Ofile.writeLine "<td style=""font-family: Times New Roman"">"  'writing new table data at column 0
				If  VarStatus = "Pass" Then
					Ofile.writeLine "<TestDescription><font color=""green"">"&VarDesc&"</font></TestDescription>"
				ElseIf VarStatus = "Fail" Then
					Ofile.writeLine "<TestDescription><font color=""black"">"&VarDesc&"</font></TestDescription>"
				Elseif VarStatus = "File" Then
					Ofile.writeLine "<TestDescription><font color=""red"">"&VarDesc&"</font></TestDescription>"
				Else
					Ofile.writeLine "<TestDescription>"&VarDesc&"</TestDescription>"
				End If
				
				Ofile.writeLine "</td>"
				Ofile.writeLine "<td style=""font-family: Times New Roman"">"  'writing new table data at column 1
				If  VarStatus = "Pass" Then
					Ofile.writeLine "<ActualResult><font color=""green"">"&VarActRes&"</font></ActualResult>"
				ElseIf VarStatus = "Fail" Then
					Ofile.writeLine "<ActualResult><font color=""red"">"&VarActRes&"</font></ActualResult>"
				ElseIf VarStatus = "File" Then
					Ofile.writeLine "<ActualResult><font color=""red"">"&VarActRes&"</font></ActualResult>"
				Else
					Ofile.writeLine "<ActualResult>"&VarActRes&"</ActualResult>"
				End If
				
				Ofile.writeLine "</td>"				
				If VarStatus = "Pass" Then  'Set TD colour to green fir test status is pass else to red
					Ofile.writeLine "<td bgcolor= ""#008000"" Style="&"""color:black"""&">"  'writing new table data at column 4
					Reporter.ReportEvent micPass,VarDesc, VarActRes
				Elseif VarStatus = "Fail" Then
					Environment.Value("intFailCount") = Environment.Value("intFailCount") + 1
					strSnapshotFile = Environment.Value("HTMLPath") & "\Snapshot\" & Environment.Value("intFailCount") & "_" & Environment.Value("strTestCase") & ".png"
					Desktop.CaptureBitmap strSnapshotFile
					strSnapshotRelativePath = "..\" & Environment.Value("HTMLFoldername") & "\Snapshot\" & Environment.Value("intFailCount") & "_" & Environment.Value("strTestCase") & ".png"
					Environment.Value("strSnapshotRelativePath") = strSnapshotRelativePath
					Ofile.writeLine "<td bgcolor= ""#FF0000"" Style="&"""color:black"""&">"  'writing new table data at column 4
					Reporter.ReportEvent micFail,VarDesc, VarActRes
				Elseif VarStatus = "File" Then					
					strFiletRelativePath = "..\" & Environment.Value("HTMLFoldername") & "\Snapshot\" & Environment.Value("strFileName")
					Ofile.writeLine "<td bgcolor= ""#FF0000"" Style="&"""color:black"""&">"  'writing new table data at column 4
				Elseif VarStatus = "Done" Then
					Ofile.writeLine "<td bgcolor= ""6495ED"" Style="&"""color:black"""&">"  'writing new table data at column 4
					Reporter.ReportEvent micDone,VarDesc, VarActRes
				Elseif VarStatus = "Stop" Then
					Ofile.writeLine "<td bgcolor= ""#FF0000"" Style="&"""color:black"""&">"  'writing new table data at column 4
				End if
				
				If  VarStatus = "Fail" Then
					Ofile.writeLine "<Status><a href="""& strSnapshotRelativePath & """ target=""_blank""><font color=""black"">"&VarStatus&"</font></a></Status>"
				ElseIf  VarStatus = "File" Then						
					Ofile.writeLine "<Status><a href="""& strFiletRelativePath & """ target=""_blank""><font color=""red"">"&VarStatus&"</font></a></Status>"
				Else
					Ofile.writeLine "<Status>"&VarStatus&"</Status>"
				End If
				
                Ofile.writeLine "</td>"
				Ofile.writeLine "<td>" 'writing new table data at column 2
				If  VarStatus = "Pass" Then
					Ofile.writeLine "<StartTime><font color=""green"">"&VarStartTime&"</font></StartTime>"
				ElseIf VarStatus = "Fail" Then
					Ofile.writeLine "<StartTime><font color=""red"">"&VarStartTime&"</font></StartTime>"
				ElseIf VarStatus = "File" Then
					Ofile.writeLine "<StartTime><font color=""red"">"&VarStartTime&"</font></StartTime>"
				Else
					Ofile.writeLine "<StartTime>"&VarStartTime&"</StartTime>"
				End If
				
				Ofile.writeLine "</td>"
				Ofile.writeLine "<td>" 'writing new table data at column 3
				If  VarStatus = "Pass" Then
					Ofile.writeLine "<EndTime><font color=""green"">"&VarEndTime&"</font></EndTime>"
				ElseIf VarStatus = "Fail" Then
					Ofile.writeLine "<EndTime><font color=""red"">"&VarEndTime&"</font></EndTime>"
				ElseIf VarStatus = "File" Then
					Ofile.writeLine "<EndTime><font color=""red"">"&VarEndTime&"</font></EndTime>"
				Else
					Ofile.writeLine "<EndTime>"&VarEndTime&"</EndTime>"
				End If
				
				Ofile.writeLine "</td>"
		End Select
		Ofile.writeLine "</tr>" 'closing row tag
	Ofile.close 'closing file
	set Ofile = Nothing  'freeing object memory for ofile and fso
	Set fso = Nothing 
End Function

'==================================================================================================================================================
' Name of the Function     			  : CloseReport
' Description       		   		 	     : This function will update the Pass Fail count in the result file
' Date and / or Version       	    : 
' Example Call							 : CloseReport()
'==================================================================================================================================================
Public Function CloseReport()

   Dim fsobj, ofile, xmlDoc, Cases, n_cases, i,  j, parentNode, statuscases, NodeToEdit
	Set fsobj = CreateObject("Scripting.FileSystemObject")  'Creating fso object
	set Ofile = fsobj.OpenTextFile(Environment.Value("strReportPath"),8) 'opens file in append mode (8)
		Ofile.writeline "</table>"  'closeing all tags
		Ofile.writeline "</body>"
		Ofile.writeline "</html>"
	Ofile.close  'close file
	set Ofile = Nothing 'freeing object memory for ofile and fso
	Set fsobj = Nothing
	
	Set xmlDoc = CreateObject("Microsoft.xMLDOM") 'opening file with xml dom object
	xmlDoc.Async = False
	xmlDoc.Load(Environment.Value("strReportPath")) 'loading html in xmdl doc

	If IsScriptTerminated = True Then
		set cases = xmlDoc.getElementsByTagName("TestDescription")
		n_cases = cases.length
		set statuscases = xmlDoc.getElementsByTagName("Status")
		statuscases = statuscases.length
		
		Set parentNode = xmlDoc.documentElement.getElementsByTagName("td")
		NodeToEdit = 10 + statuscases *  4 + n_cases - 1
		parentNode.item(NodeToEdit).setattribute "style","color:red"
		
		Set parentNode = xmlDoc.SelectSingleNode("/html/body/table[2]/tr["&n_cases &"]/td[0]")
		parentNode.Text = parentNode.Text & gstrAppendText
	End If
	
	Set parentNode = xmlDoc.SelectSingleNode("/html/body/table[0]/tr[0]/td[1]")  'writing Total count in html
		parentNode.Text = Environment.Value("TESTSCENARIOCOUNT") 'TotalTestScnExec
	Set parentNode = xmlDoc.SelectSingleNode("/html/body/table[0]/tr[1]/td[1]")  'writing Pass Count in html
		parentNode.Text = Environment.Value("TESTSCENARIOCOUNT") - Environment.Value("TOTALFAILCOUNT") 'PassCnt
	Set parentNode = xmlDoc.SelectSingleNode("/html/body/table[0]/tr[2]/td[1]")  'writing Fail Count in html
		parentNode.Text = Environment.Value("TOTALFAILCOUNT") 'TotalTestScnExec - PassCnt
	Set parentNode = xmlDoc.SelectSingleNode("/html/body/table[0]/tr[4]/td[1]")  'writing end time
		parentNode.Text = Now
	xmlDoc.Save(Environment.Value("strReportPath")) 'saving file
	UpdateLog "Test Execution completed"
End Function
'==================================================================================================================================================
' Name of the Function			: CreateLog
' Description       		   	: This function should be called in the beginning of the execution in order to create the log file
' Date and / or Version       	 	: 
' Example Call				: CreateLog()
'==================================================================================================================================================
Public Function CreateLog(strTCName)

			DateTime = Date &"_" &Time
			Set regEx = New RegExp
			regEx.global = true
			regEx.pattern = "[/\ \:  ]"
			TimeStamp = regEx.replace(DateTime, "_")
			
			strResultPathSplit = Split(Environment.Value("TestDir"), "\")
			strFolderLen = Len(strResultPathSplit(UBound(strResultPathSplit)))
			strTotalLen = Len(Environment.Value("TestDir"))
			strLogPath = Left(Environment.Value("TestDir"), (strTotalLen-strFolderLen))
			
			Environment.Value("strLogPath") = strLogPath & strTCName & "_log_"&TimeStamp&".log"
			
			Set fsolog = CreateObject("Scripting.FileSystemObject") 	'Creating fso object
			Set ofilelog=fsolog.CreateTextFile(Environment.Value("strLogPath"))
			ofilelog.WriteLine String(80,"-")
			ofilelog.WriteLine "QTP Execution Log File"
			ofilelog.WriteLine String(80,"-")
			ofilelog.WriteLine "DATE " & space(10) & "TIME" & space(10) & " LOG DESCRIPTION"
			ofilelog.WriteLine String(80,"-")

End Function

'==================================================================================================================================================
' Name of the Function              : UpdateLog
' Description       		    : This function is used to update the TESTCASE or TESTSTEP result
' Date and / or Version       	    : 
' Example Call	  		    : UpdateLog "Test Sequence", "", "Action", "Step Description", "Pass"
'==================================================================================================================================================

Public Function UpdateLog(strDetail)
			Dim fso, ofile
			Set fso = CreateObject("Scripting.FileSystemObject") 	'Creating fso object
			Set ofile = fso.OpenTextFile(Environment.Value("strLogPath"),8) 'Opening file in append mode (8). appends lines at end of file.
'			ofilelog.WriteLine Now & space(5) & strDetail
			ofile.WriteLine Now & space(5) & strDetail
			
			Set ofile = Nothing  'freeing object memory for ofile and fso
			Set fso = Nothing 

End Function
'==================================================================================================================================================
' Name of the Function     	   :CloseLog
' Description       		   : This function is used to Close the Log file
' Date and / or Version       	   : 
' Example Call			   : CloseLog()
'==================================================================================================================================================
Public Function CloseLog()
			Dim fso, ofile
			Set fso = CreateObject("Scripting.FileSystemObject") 	'Creating fso object
			Set ofile = fso.OpenTextFile(Environment.Value("strLogPath"),8) 'Opening file in append mode (8). appends lines at end of file.
			ofile.WriteLine String(80,"-")
			ofile.close 'close the file
			Set ofile = Nothing  'Releasing object memory for ofile and fso
			Set fsolog = Nothing 
End Function

'========================================================================================================= 
'	Name of the Function		   : Execution_log
'	Scope of the Function  		   : Public
'	Author 		   		   : Shrinidhi Holla
'	Description 			   : This function is To capture the execution results.
'	Parameters accepted		   : a) <strScenarioName> Name of the Scenario
'					   : b) <strTestCaseID> Test Case ID
'					   : c) <strTestCaseDesc> Description of the Test Case under Execution
'					   : d) <strResults> Result of the test Under Execution						
'	Parameter returned 		   : Null
'	Date and / or Version 	: 12-01-2010,version 1.0
'=========================================================================================================
Public Function Execution_log(strScenarioName,strTestCaseID, strTestCaseDesc, strResults)

        Dim coulmn
		Dim FSO
		'Open Result log file
		Set  strAppRes = CreateObject( "Excel.Application") 
		strAppRes.WorkBooks.Open Environment.Value("RESULT_FILE")
		Set strResultSheet = strAppRes.ActiveWorkbook 
		Set strResultExcel = strResultSheet.Worksheets(1)

		'Set the row and column To append
		intRow = Environment.Value("RESULT_FILE_ROW") + 1
		Environment.Value("RESULT_FILE_ROW") = intRow
		intCol = 1
       
		'Enter scenario name or Test case Id, Description and Results
		If strScenarioName <> "" Then
			strResultExcel.Cells(intRow, intCol).Font.Bold = "True"
			strResultExcel.Cells(intRow, intCol) = strScenarioName 
		Else
			 strResultExcel.Cells(intRow, intCol + 1) = strTestCaseID 
			 strResultExcel.Cells(intRow, intCol + 2).Font.Bold = "True"
			 Environment.Value("TESTCASE_COUNT") = Environment.Value("TESTCASE_COUNT") + 1
			If(Ucase(strResults) <> "PASS")Then
				Environment.Value("FAIL_COUNT") = Environment.Value("FAIL_COUNT") + 1
				strResultExcel.Cells(intRow, intCol + 2).Font.ColorIndex = 1 
				strResultExcel.Cells(intRow, intCol + 2).Interior.ColorIndex = 3 
				strResultExcel.Cells(intRow, intCol + 2) = "FAIL"   
			Else
				Environment.Value("PASS_COUNT") = Environment.Value("PASS_COUNT") + 1
				strResultExcel.Cells(intRow, intCol + 2).Interior.ColorIndex = 4 
				strResultExcel.Cells(intRow, intCol + 2) = "PASS"   
			End If
		End If
	 
		strResultSheet.Save
		'strResultSheet.SaveAs strResultFile
		Set strResultSheet = Nothing
		Set strResultExcel = Nothing
		strAppRes.Quit

End Function

'========================================================================================================= 
'	Name of the Function		   : CreateExecutionSummaryFile
'	Scope of the Function  		   : Public
'	Author 		   		   : Shrinidhi Holla
'	Description 			   : This function is used To Create the Execution Log file.
'	Parameters accepted		   : None
'	Parameter returned 		   : Null
'	Date and / or Version 	: 12-01-2010,version 1.0
'=========================================================================================================
Public Function CreateExecutionSummaryFile()

    'File name string formation
	strCurrentDate = Date
	strDate = Split(strCurrentDate,"/" ,-1,1)
	strFileName = "AutomationExcecutionResults_"&strDate(0) &strDate(1)&strDate(2) & ".xls"

	strFinalPath = Environment.Value("ResultPath") & strFileName
	
	strResultFile = strFinalPath
	Environment.Value("RESULT_FILE") = strResultFile
	Environment.Value("RESULT_FILE_ROW") = 1

	''Create Results folder if not present
	Set fso1= CreateObject("Scripting.FileSystemObject")
	If (fso1.FileExists(strResultFile))Then
		fso1.DeleteFile(strResultFile)
	End If

	'Create Heading row
	row= 1
	Set  strAppRes = CreateObject( "Excel.Application") 
	strAppRes.WorkBooks.Add
	Set strResultSheet = strAppRes.ActiveWorkbook 
	Set strResultExcel = strResultSheet.Worksheets(1)

	For intColCount = 1 To 4 
		strResultExcel.Cells(row,intColCount).Font.Bold = "True"
		strResultExcel.Cells(row,intColCount).Font.ColorIndex = 1 
		strResultExcel.Cells(row,intColCount).Interior.ColorIndex = 15
	Next
	
	coulmn=0
	strResultExcel.Cells(row,coulmn+1) ="Scenario Name"
	strResultExcel.Cells(row,coulmn+2) ="Test Case Name"
	strResultExcel.Cells(row,coulmn+3) ="Results"

	strResultSheet.SaveAs(strResultFile)
	wait(3)
	strAppRes.Quit
	Set strResultSheet = Nothing
	Set strResultExcel = Nothing
	Set fso1 = Nothing

	Environment.Value("TESTCASE_COUNT") = 0
	Environment.Value("FAIL_COUNT") = 0
	Environment.Value("PASS_COUNT") = 0

End Function


'========================================================================================================= 
'	Name of the Function		   : LoadRepository
'	Scope of the Function  		   : Public
'	Author 		   		   : Shrinidhi Holla
'	Description 			   : This function is attach the Object Repository
'	Parameters accepted		   : ORName
'	Parameter returned 		   	   : Null
'	Date and / or Version 	                    : 15-Mar-2011,version 1.0
'=========================================================================================================
Function LoadRepository(strRepPathName)
    strObjPath = Environment.Value("Resource_Path")
	Set qtApp = CreateObject("QuickTest.Application") ' Create the Application object
	strRepPathName = Split(strRepPathName,";",-1,1)
	For intCounter = 0 To UBound(strRepPathName)
		Set qtReposiTories = qtApp.Test.Actions("Action1").ObjectReposiTories ' Get the object reposiTories collection object of the "Login" action
		strObjRepPath = strObjPath & strRepPathName(intCounter)
		If qtReposiTories.Find(strObjRepPath) = -1 Then ' If the reposiTory cannot be found in the collection
			qtReposiTories.Add strObjRepPath, 1 ' Add the reposiTory To the collection
		End If
	Next
	Set qtApp = Nothing
	Set qtReposiTories = Nothing
End Function

'========================================================================================================= 
'	Name of the Function		   : CloseExecutionSummary
'	Scope of the Function  		   : Public
'	Author 		   		   : Shrinidhi Holla
'	Description 			   : This function is log the Execution Details of the Test and Inserts Chart in Excel Result.
'	Parameters accepted		   : None
'	Parameter returned 		   	   : Null
'	Date and / or Version 	                    : 12-01-2010,version 1.0
'=========================================================================================================
Public Function CloseExecutionSummary()

		Dim coulmn
		Dim FSO
		'Open Result log file
		Set  strAppRes = CreateObject( "Excel.Application")
		strAppRes.WorkBooks.Open Environment.Value("RESULT_FILE")
		Set strResultSheet = strAppRes.ActiveWorkbook
		Set strResultExcel = strResultSheet.Worksheets(1)
	
		'Set the row and column To append
		intRow = Environment.Value("RESULT_FILE_ROW") + 2
		intCol = 1
	
		strAppRes.Cells(intRow, intCol).Font.Bold = "True"
		strAppRes.Cells(intRow, intCol) = "EXECUTION SUMMARY"
		strAppRes.Cells(intRow, intCol+1) = Now
		
		strAppRes.Cells(intRow+1, intCol) = "Test Cases Executed"
		strAppRes.Cells(intRow+1, intCol+1) = Environment.Value("TESTCASE_COUNT")
		strAppRes.Cells(intRow+2, intCol) = "Test Cases Passed" 
		strAppRes.Cells(intRow+2, intCol+1) = Environment.Value("PASS_COUNT")
		strAppRes.Cells(intRow+3, intCol) = "Test Cases Failed" 
		strAppRes.Cells(intRow+3, intCol+1) = Environment.Value("FAIL_COUNT")
	
		strColumnName = Chr(intCol+1+64)
		strFromRange =  strColumnName&(intRow+2)
		strToRange =  strColumnName&(intRow+3)

		Set rng = strAppRes.Sheets("Sheet1").Range(strFromRange,strToRange)
		
		strAppRes.Charts.Add
		strAppRes.ActiveChart.ChartType = 5
		strAppRes.ActiveChart.HasLegend = False
		strAppRes.ActiveChart.SeriesCollection.NewSeries
		
        strAppRes.ActiveChart.SeriesCollection(1).HasDataLabels= True
        strAppRes.ActiveChart.SeriesCollection(1).DataLabels.Position = 3
		
		strAppRes.ActiveChart.SetSourceData rng, 2
		strAppRes.ActiveChart.Location 2, "Sheet1"
		strAppRes.ActiveChart.SeriesCollection(1).Name = "Automation Test Coverage"
        
		'strAppRes.Save                                ' Commented by Shrinidhi to overcome the popup issue,
        strAppRes.ActiveWorkbook.Save
		strAppRes.quit
End Function

'==================================================================================================================================================
' Name of the Function     			  : fnc_wait
' Description       		   		 	     : This function is used to wait until the object to appear in the screen
' Date and / or Version       	    : 
' Example Call							 : OperateOnWebEdit("Browser("TIPS").Page("TitleGrid").WebEdit"Login")
'==================================================================================================================================================
Function fnc_wait(v_ob_wobj)
		Dim v_bl_Flag,iCounter

		'Modified during Demo APP poc - Abdul
		iCounter=1
		Do
			iCounter=iCounter+1
			wait(1)
			v_bl_Flag=v_ob_wobj.Exist(1)
		LOOP WHILE (Not v_bl_Flag and iCounter<=10)

		fnc_wait = v_bl_Flag

End Function

'==================================================================================================================================================
' Name of the Function     			  : OperateOnWebEdit
' Description       		   		 	     : This function is used to enter or retrieve the values on / from the WebEdit object
' Date and / or Version       	    : 
' Example Call							 : OperateOnWebEdit("Browser(""TIPS"").Page(""TitleGrid"")","LoginName", "SETVALUE", "Test123")
'==================================================================================================================================================
Function OperateOnWebEdit(strObjectHierarchy, strObjectName, strAction_Name, strData)
   
		If  strAction_Name <> "STOREVALUE" AND Instr(strData, "VAR_") = 1 Then				' If the value to be taken from already saved variable
			strData = Environment.Value(strData)
		End If

		If  Instr(strData, "fn_") = 1 Then
			strData = Environment.Value(strData)
		End If

		Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.WebEdit(strObjectName)
		StepStartTime = Time
        If fnc_wait(ActualObject) = "True" Then
				Select Case strAction_Name
				
				Case "SETVALUE"
				
					Wait(1)
					ObjectHierarchy.WebEdit(strObjectName).Set strData							
					If Instr(Environment.Value("strDescription"), "Password") > 0 Then
						UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime, Time, "Value: <font color=""blue""> <b><i>" & strData & "</i></b></font> is entered " & " in <b>'" & Mid(strObjectName,4,Len(strObjectName)-3) & "'</b> field successfully", "Done"
						'UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime, Time, "The <font color=""blue""> <b><i> Password </i></b> </font>is entered in field <b>'"  & Mid(strObjectName,4,Len(strObjectName)-3) & "'</b>", "Done"
					Else
						UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime, Time, "Value: <font color=""blue""> <b><i>" & strData & "</i></b></font> is entered " & " in <b>'" & Mid(strObjectName,4,Len(strObjectName)-3) & "'</b> field successfully", "Done"
					End If
					
				Case "CLICK"
				
					   ObjectHierarchy.WebEdit(strObjectName).Click									
					   UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime, Time, "The WebEdit: <b>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</b> is clicked successfully", "Done"
					
				Case "TYPEVALUE"
				
					  ObjectHierarchy.WebEdit(strObjectName).Type strData
					  UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime, Time, "The Value: <font color=""blue""> </i></b>" & strData & "</i></b> </font>is entered in field <b>'" & Mid(strObjectName,4,Len(strObjectName)-3) & "'</b>", "Done"
					  
				Case "COMPAREVALUE"
				
					If Instr(strData,"fn_") >0 or instr(strData,"VAR_") >0  Then
						strExp = Environment.Value(strData)
					Else
						strExp = strData
					End If	
					
					strActual = ObjectHierarchy.WebEdit(strObjectName).GetROProperty("Value")
					If Trim(UCASE(strExp)) = Trim(UCASE(strActual)) Then
						UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Expected Value - "& strExp &" and Actual value - "& strActual & " are matching for field <b>'" & Mid(strObjectName,4,Len(strObjectName)-3) & "</b></font>", "Pass"
					Else
						UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Value mismatch, Expected Value is: <b><i>" & strExp & "</b></i>, and Actual value on the application is: <b><i>" & strActual & "</b></i></font>", "Fail"
						Environment.Value("TestStepLog") = "False"
					End If
					
				Case "CHECKEXIST"
				
					strExp = strData
					strActual = ObjectHierarchy.WebEdit(strObjectName).Exist
					If UCASE(strExp) = UCASE(CStr(strActual)) and UCASE(strExp) = "TRUE" Then
						UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The field <b>'" & Mid(strObjectName,4,Len(strObjectName)-3) & "'</b> displayed successfully</font>", "Pass"
					ElseIf UCASE(strExp) <> UCASE(CStr(strActual)) and UCASE(strExp) = "TRUE" Then
						UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The field <b>'" & Mid(strObjectName,4,Len(strObjectName)-3) & "'</b> does not exist</font>", "Fail"
						Environment.Value("TestStepLog") = "False"
					ElseIf UCASE(strExp) = UCASE(CStr(strActual)) and UCASE(strExp) = "FALSE" Then
						UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The field <b>'" & Mid(strObjectName,4,Len(strObjectName)-3) & "'</b> not displayed as expected</font>", "Pass"
					ElseIf UCASE(strExp) <> UCASE(CStr(strActual)) and UCASE(strExp) = "FALSE" Then
						UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object exists</font>", "Fail"
						Environment.Value("TestStepLog") = "False"
				  	End If
				  	
				Case "SETPASSWORD"
				
					 ObjectHierarchy.WebEdit(strObjectName).SetSecure strData
					 UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime, Time, "The Password is entered", "Done"
					 
				Case "STOREVALUE"
				
					 strVariableName = strData
					 Environment.Value("VAR_"& strVariableName) = ObjectHierarchy.WebEdit(strObjectName).GetROProperty("value")
					 UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The <font color=""blue""> <b><i>"& strData &"</i><b></font> value: <font color=""blue""> <b><i>"& Environment.Value("VAR_"& strVariableName)& "</i></b></font> is stored in field ", "Done"
					 
				Case "CHECKENABLED"
				
					blnObjDisable= ObjectHierarchy.WebEdit(strObjectName).GetROProperty("disabled")
					If blnObjDisable = "1" AND UCASE(strData) = "FALSE" Then
						UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Disabled as expected</font>", "Pass"
					ElseIf blnObjDisable = "0" AND UCASE(strData) = "TRUE" Then
						UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Enabled as expected</font>", "Pass"
					ElseIf blnObjDisable = "0" AND UCASE(strData) = "FALSE" Then
						UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Enabled</font>", "Fail"
						Environment.Value("TestStepLog") = "False"
					ElseIf blnObjDisable = "1" AND UCASE(strData) = "TRUE" Then
						UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Disabled</font>", "Fail"
						Environment.Value("TestStepLog") = "False"
					End If
					
				End Select
							
        Else
                UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The WebEdit - <b>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If

End Function

'==================================================================================================================================================
' Name of the Function     			  : OperateOnWebButton
' Description       		   		 	     : This function is used to enter or retrieve the values on / from the Web object
' Date and / or Version       	    : 
' Example Call							 : OperateOnWebButton("Browser(""TIPS"").Page(""TitleGrid"")","OK", "CLICK", "")
'==================================================================================================================================================
Function OperateOnWebButton(strObjectHierarchy, strObjectName, strAction_Name, strData)
   
		Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.WebButton(strObjectName)
		StepStartTime = Time
		Call fnc_wait(ActualObject)
		
		If fnc_wait(ActualObject) = "True" Then
					Select Case strAction_Name
					  Case "CLICK"					
							  ObjectHierarchy.WebButton(strObjectName).Click
							  UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime, Time, "The Button: <b>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</b> is clicked", "Done"
					  Case "CHECKEXIST"
							  strExp = strData
							  strAct = ActualObject.Exist(0)
							  If UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The button <b>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</b> displayed succesfully</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The button <b>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</b> does not exist</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  ElseIf UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The button <b>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</b> is not displayed as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The button <b>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</b> is displayed</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
						Case "CHECKENABLED"
							  blnObjDisable= ActualObject.GetROProperty("disabled")
								If blnObjDisable = "1" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Button <b>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</b> is Disabled as expected</font>", "Pass"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Button <b>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</b> is Enabled as expected</font>", "Pass"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The button <b>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</b> is Enabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								ElseIf blnObjDisable = "1" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The button <b>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</b> is Disabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								End If
							End Select
								
	           Else
							UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Button - <b>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</b> does not exist</font>", "Fail"
							Environment.Value("TestObjectFlag") = "False"
			End If		
End Function

'==================================================================================================================================================
' Name of the Function     			  : OperateOnWebList
' Description       		   		 	     : This function is used to select or retrieve the values from the WebList object
' Date and / or Version       	    : 
' Example Call							 : OperateOnWebList("Browser(""TIPS"").Page(""TitleGrid"")","StatusCode", "SELECTVALUE", "Test")
'==================================================================================================================================================
Function OperateOnWebList(strObjectHierarchy, strObjectName, strAction_Name, strData)

Wait(2)
		If  strAction_Name <> "STOREVALUE" AND Instr(strData, "VAR_") = 1 Then				' If the value to be taken from already saved variable
				strData = Trim(Environment.Value(strData))
		End If

	    If  Instr(strData, "fn_") = 1Then
			strData = Environment.Value(strData)
		End If
			
	    Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.WebList(strObjectName)
		StepStartTime = Time
        If fnc_wait(ActualObject) = "True" Then
				Select Case strAction_Name
					  Case "SELECTVALUE"
							  On Error Resume Next
							   Err.Clear
							  ObjectHierarchy.WebList(strObjectName).Select strData							 
							  If Err.Number <> 0 Then
  									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value <font color=""blue""> <b><i>" & trim(strData) & " </i></b></font> is not available in the Dropdown <b>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</b>" , "Fail"
									Environment.Value("TestStepLog") = "False"
									Environment.Value("TestObjectFlag") = "False"
							  Else
									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value <font color=""blue""> <b><i>" &  trim(strData) & "</i></b></font> is selected from the Dropdown", "Done"
							  End If
							  
								Err.Clear
					  Case "COMPAREVALUE"
						  '********************* Updated by Basavaraj 14 Feb, 2012************************************
						  If instr(strData,"fn_") >0 or instr(strData,"VAR_") >0  Then
								strExp = Environment.Value(strData)
						  Else
								strExp = strData
						  End If							  
						  '********************* Updated by Basavaraj 14 Feb, 2012************************************
							  strActual = ObjectHierarchy.WebList(strObjectName).GetROProperty("value")
							  If Trim(UCASE(strExp)) = Trim(UCASE(strActual)) Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Expected Value - "& strExp &" and Actual value - "& strActual & " are matching</font>", "Pass"
							  Else
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Value mismatch, Actual value is: <i>" & strActual & "</i></font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
							  
					  Case "STOREVALUE"
							 strVariableName = strData
							 Environment.Value("VAR_"&strVariableName) = ObjectHierarchy.WebList(strObjectName).GetROProperty("value")	
							 UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The <font color=""blue""><b><i>"& strData &"<i><b></font> value is stored</font>", "Done"
							 
					  Case "CHECKEXIST"
							  strExp = strData
							  strAct = ObjectHierarchy.WebList(strObjectName).Exist
							  If UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object exists as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object does not exist</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  ElseIf UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object does not exist as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object exists</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
						Case "CHECKENABLED"
								blnObjDisable= ObjectHierarchy.WebList(strObjectName).GetROProperty("disabled")
								'**Start***Updated by Manish on 6/29/11***									
								If blnObjDisable = "1" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Disabled as expected</font>", "Pass"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Enabled as expected</font>", "Pass"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Enabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								ElseIf blnObjDisable = "1" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Disabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								End If
						Case "CHECKVALUEINDROPDOWN"			
								arrExpItem = Split(strData,",",-1,1)
                                intItemCount = ActualObject.GetROProperty("items count")
                                strAllItem =  ActualObject.GetROProperty("all items")
                                arrActItem = Split(strAllItem,";",-1,1)
                                For intExpItem = 0 to Ubound(arrExpItem)  
                                    For intActItem = 0 to Ubound(arrActItem) 
                                        If instr(Ucase(Trim(arrActItem(intActItem))),Ucase(Trim(arrExpItem(intExpItem)))) Then
                                            isItemFound = True
                                            Exit For
                                        Else
                                            isItemFound = False       
                                        End If
                                    Next
                                    If  isItemFound = True Then
                                        UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green""><b>"& arrExpItem(intExpItem)  & " </b>: is available in DropDown</font>", "Pass"
                                    Else
                                        UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red""><b>"& arrExpItem(intExpItem) & "</b> : Not Available in DropDown</font>", "Fail"
                                        Environment.Value("TestStepLog") = "False"                     
                                    End If
                                Next
						Case "CHECKVALUENOTINDROPDOWN"						' Added by Shrinidhi on 21-Nov-2011
								blnValueFound = "True"
								strAllItems = ObjectHierarchy.WebList(strObjectName).GetROProperty("all items")
								strAllItemsSplit = Split (strAllItems,";")
								For intCount = 0 to UBound(strAllItemsSplit)
										If  UCASE(Trim(strAllItemsSplit(intCount))) = UCASE(Trim(strData)) Then
											UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red""><b>"& strData & "</b> : is Available in DropDown</font>", "Fail"
											blnValueFound = "True"
											Exit For
										Else
											blnValueFound = "False"
										End If
								Next
								If  blnValueFound = "False" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green""><b>"& strData  & " </b>: is not available in DropDown</font>", "Pass"
										Environment.Value("TestStepLog") = "False"
								End If
						Case "GETITEMSCOUNT"
							 strVariableName = strData
							 strCount =  ObjectHierarchy.WebList(strObjectName).GetROProperty("items count") -1
							 Environment.Value("VAR_"&strVariableName) = strCount					
							 UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The Total items count is  :"& strCount &" value is stored</font>", "Done"

				End Select
	   Else
                UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The WebList - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If
End Function
'==================================================================================================================================================
' Name of the Function     			  : OperateOnWinList
' Description       		   		 	     : This function is used to select or retrieve the values from the WinList object
' Date and / or Version       	    : 
' Example Call							 : OperateOnWinList(strObjectHierarchy, strObjectName, strAction_Name, strData)
'==================================================================================================================================================
Function OperateOnWinList(strObjectHierarchy, strObjectName, strAction_Name, strData)

		If  strAction_Name <> "STOREVALUE" and Instr(strData, "VAR_") = 1 Then				' If the value to be taken from already saved variable
				strData = Trim(Environment.Value(strData))
		End If

	    If  Instr(strData, "fn_") = 1 Then
			strData = Environment.Value(strData)
		End If
		
		strResData = Replace(strData,"/","")
		strResData = Replace(strResData,"-","")
		
	    Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.WinList(strObjectName)
		
		StepStartTime = Time
        If fnc_wait(ActualObject) = "True" Then
				Select Case strAction_Name
					  Case "SELECTVALUE"
							  On Error Resume Next
							  ObjectHierarchy.WinList(strObjectName).Select strData
							  strData = Trim(strData)
							  strAppData = ObjectHierarchy.WinList(strObjectName).GetROProperty("Selection")							  

							  If StrComp(Trim(strAppData),strData) = 0 Then
							  		ObjectHierarchy.WinList(strObjectName).Activate strAppData
									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value: <font color=""blue""><b><i>" & strAppData & "</i></b></font> is selected from the Dropdown <b>"  &Mid(strObjectName,4,Len(strObjectName)-3)& "</b>", "Done"
							  elseIf StrComp(strAppData,strData) <> 0 Then
  									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value: <font color=""blue""><b><i></font>" & strAppData & "</i></b></font> is not available in the Dropdown <b>"  &Mid(strObjectName,4,Len(strObjectName)-3) &"</b>", "Fail"
									Environment.Value("TestStepLog") = "False"
									Environment.Value("TestObjectFlag") = "False"
							  End If
							  
					  Case "COMPAREVALUE"
						  If instr(strData,"fn_") >0 or Instr(strData,"VAR_") >0  Then
								strExp = Environment.Value(strData)
						  Else
								strExp = strData
						  End If
								
							  ObjectHierarchy.WinList(strObjectName).Select strData								
							  strActual = Trim(ObjectHierarchy.WinList(strObjectName).GetROProperty("Selection"))
							  If StrComp(strExp,strActual) = 0 Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Expected Value - "& strExp &" and Actual value - "& strActual & " are matching</font>", "Pass"
							  Else
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Value mismatch, Actual value is: <i>" & strActual & "</i></font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
							  
					  Case "CHECKEXIST"
							  strExp = strData
							  strAct = ObjectHierarchy.WinList(strObjectName).Exist(5)
							  If UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object exists as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object does not exist</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  ElseIf UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object does not exist as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object exists</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
					 Case "COMPAREANDSELECT"
							  strDataVal = fn_ListDataEncapsulation(strScenarioName,strDataVal)
							  
								If strDataVal <> "" then
									  	  ObjectHierarchy.WinList(strObjectName).Select strDataVal
										  strAppData = ObjectHierarchy.WinList(strObjectName).GetROProperty("Selection")							  
								
										  If StrComp(Trim(strAppData),Trim(strDataVal)) = 0 Then
										  		ObjectHierarchy.WinList(strObjectName).Activate strAppData
												UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value: <b><i>" & strAppData & "</i></b> is selected from the Dropdown", "Done"
										  elseIf StrComp(strAppData,strData) <> 0 Then
												UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value: <b><i>" & strAppData & "</i></b> is not available in the Dropdown", "Fail"
												Environment.Value("TestStepLog") = "False"
												Environment.Value("TestObjectFlag") = "False"
										  End If
								else
												UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value: <b><i>" & strDataVal & "</i></b> is not available in the Dropdown", "Fail"
												Environment.Value("TestStepLog") = "False"
												Environment.Value("TestObjectFlag") = "False"
								End if
								
					 Case "CHECKVALUENOTPRESENTINLIST"
					 
							If instr(strData,"fn_") >0 or Instr(strData,"VAR_") >0  Then
								strExp = Environment.Value(strData)
							Else
								strExp = strData
							End If
													
							strAllItems = Trim(ObjectHierarchy.WinList(strObjectName).GetROProperty("all items"))
							If Instr(strAllItems,strExp) = 0 Then
								UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">"& strExp &" not present in the list </font>", "Pass"
							Else
								UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red""> " & strExp & " present in the list </font>", "Fail"
								Environment.Value("TestStepLog") = "False"
							End If
							
					Case "CHECKVALUEPRESENTINLIST"
						
							If instr(strData,"fn_") >0 or Instr(strData,"VAR_") >0  Then
								strExp = Environment.Value(strData)
							Else
								strExp = strData
							End If
													
							strAllItems = Trim(ObjectHierarchy.WinList(strObjectName).GetROProperty("all items"))
							If Instr(strAllItems,strExp) > 1 Then
								UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">"& strExp &" present in the list </font>", "Pass"
							Else
								UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red""> " & strExp & " not present in the list </font>", "Fail"
								Environment.Value("TestStepLog") = "False"
							End If
							
				End Select
	   Else
                UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The WinList - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If
		
End Function
'==================================================================================================================================================
' Name of the Function     			  : OperateOnPbList
' Description       		   		  : This function is used to select or retrieve the values from the WinList object
' Date and / or Version       	      : 
' Example Call						  : OperateOnPbList(strObjectHierarchy, strObjectName, strAction_Name, strData)
'==================================================================================================================================================
Function OperateOnPbList(strObjectHierarchy, strObjectName, strAction_Name, strData)

		If  strAction_Name <> "STOREVALUE" and Instr(strData, "VAR_") = 1 Then				' If the value to be taken from already saved variable
				strData = Trim(Environment.Value(strData))
		End If

	    If  Instr(strData, "fn_") = 1 Then
			strData = Environment.Value(strData)
		End If
		
		strResData = Replace(strData,"/","")
		strResData = Replace(strResData,"-","")
		
	    Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.PbList(strObjectName)
		
		StepStartTime = Time
        If fnc_wait(ActualObject) = "True" Then
				Select Case strAction_Name
					  Case "SELECTVALUE"
							  On Error Resume Next
							  ObjectHierarchy.PbList(strObjectName).Select strData
							  strData = Trim(strData)
							  strAppData = ObjectHierarchy.PbList(strObjectName).GetROProperty("Selection")							  

							  If StrComp(Trim(strAppData),strData) = 0 Then
							  		ObjectHierarchy.PbList(strObjectName).Activate strAppData
									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value: <font color=""blue""><b><i>" & strAppData & "</i></b></font> is selected from the Dropdown <b>"  &Mid(strObjectName,4,Len(strObjectName)-3)& "</b>", "Done"
							  elseIf StrComp(strAppData,strData) <> 0 Then
  									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value: <font color=""blue""><b><i></font>" & strAppData & "</i></b></font> is not available in the Dropdown <b>"  &Mid(strObjectName,4,Len(strObjectName)-3) &"</b>", "Fail"
									Environment.Value("TestStepLog") = "False"
									Environment.Value("TestObjectFlag") = "False"
							  End If
							  
					  Case "COMPAREVALUE"
						  If instr(strData,"fn_") >0 or Instr(strData,"VAR_") >0  Then
								strExp = Environment.Value(strData)
						  Else
								strExp = strData
						  End If
								
							  ObjectHierarchy.PbList(strObjectName).Select strData								
							  strActual = Trim(ObjectHierarchy.PbList(strObjectName).GetROProperty("Selection"))
							  If StrComp(strExp,strActual) = 0 Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Expected Value - "& strExp &" and Actual value - "& strActual & " are matching</font>", "Pass"
							  Else
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Value mismatch, Actual value is: <i>" & strActual & "</i></font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
							  
					  Case "CHECKEXIST"
							  strExp = strData
							  strAct = ObjectHierarchy.PbList(strObjectName).Exist(5)
							  If UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object exists as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object does not exist</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  ElseIf UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object does not exist as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object exists</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
					 Case "COMPAREANDSELECT"
							  strDataVal = fn_ListDataEncapsulation(strScenarioName,strDataVal)
							  
								If strDataVal <> "" then
									  	  ObjectHierarchy.PbList(strObjectName).Select strDataVal
										  strAppData = ObjectHierarchy.PbList(strObjectName).GetROProperty("Selection")							  
								
										  If StrComp(Trim(strAppData),Trim(strDataVal)) = 0 Then
										  		ObjectHierarchy.PbList(strObjectName).Activate strAppData
												UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value: <b><i>" & strAppData & "</i></b> is selected from the Dropdown", "Done"
										  elseIf StrComp(strAppData,strData) <> 0 Then
												UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value: <b><i>" & strAppData & "</i></b> is not available in the Dropdown", "Fail"
												Environment.Value("TestStepLog") = "False"
												Environment.Value("TestObjectFlag") = "False"
										  End If
								else
												UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value: <b><i>" & strDataVal & "</i></b> is not available in the Dropdown", "Fail"
												Environment.Value("TestStepLog") = "False"
												Environment.Value("TestObjectFlag") = "False"
								End if
								
					 Case "CHECKVALUENOTPRESENTINLIST"
					 
							If instr(strData,"fn_") >0 or Instr(strData,"VAR_") >0  Then
								strExp = Environment.Value(strData)
							Else
								strExp = strData
							End If
													
							strAllItems = Trim(ObjectHierarchy.PbList(strObjectName).GetROProperty("all items"))
							If Instr(strAllItems,strExp) = 0 Then
								UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">"& strExp &" not present in the list </font>", "Pass"
							Else
								UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red""> " & strExp & " present in the list </font>", "Fail"
								Environment.Value("TestStepLog") = "False"
							End If
							
					Case "CHECKVALUEPRESENTINLIST"
						
							If instr(strData,"fn_") >0 or Instr(strData,"VAR_") >0  Then
								strExp = Environment.Value(strData)
							Else
								strExp = strData
							End If
													
							strAllItems = Trim(ObjectHierarchy.PbList(strObjectName).GetROProperty("all items"))
							If Instr(strAllItems,strExp) > 1 Then
								UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">"& strExp &" present in the list </font>", "Pass"
							Else
								UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red""> " & strExp & " not present in the list </font>", "Fail"
								Environment.Value("TestStepLog") = "False"
							End If
							
				End Select
	   Else
                UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The WinList - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If
		
End Function
'==================================================================================================================================================
' Name of the Function     			  : OperateOnLink
' Description       		   		 	     : This function is used to click on the Link
' Date and / or Version       	    : 
' Example Call							 : OperateOnLink("Browser(""TIPS"").Page(""TitleGrid"")","GoTo", "CLICK", "")
'==================================================================================================================================================
Function OperateOnLink(strObjectHierarchy, strObjectName, strAction_Name, strData)

	    Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.Link(strObjectName)
		StepStartTime = Time
		Call fnc_wait(ObjectHierarchy)
		If fnc_wait(ObjectHierarchy) = "True" Then
				Select Case strAction_Name
							  Case "CLICK"
							  Wait(1)
								'ObjectHierarchy.Link(strObjectName).FireEvent("onmouseover")
							  ObjectHierarchy.Link(strObjectName).Click
							  UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Link " & Mid(strObjectName,4,Len(strObjectName)-3) & " is Clicked successfully", "Done"
						  Case "CHECKEXIST"
							  strExp = strData
							  strAct = ObjectHierarchy.Link(strObjectName).Exist(10)
							  If UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object exists as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object does not exist</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  ElseIf UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object does not exist as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object exists</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
							  Case "FireEventOnMouseOver"
									  ObjectHierarchy.Link(strObjectName).FireEvent("onmouseover")
									  UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Link is Clicked", "Done"
							  Case "CHECKENABLED"
										blnObjDisable= ObjectHierarchy.Link(strObjectName).GetROProperty("disabled")							
										If blnObjDisable = "1" AND UCASE(strData) = "FALSE" Then
												UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Disabled as expected</font>", "Pass"
										ElseIf blnObjDisable = "0" AND UCASE(strData) = "TRUE" Then
												UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Enabled as expected</font>", "Pass"
										ElseIf blnObjDisable = "0" AND UCASE(strData) = "FALSE" Then
												UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Enabled</font>", "Fail"
												Environment.Value("TestStepLog") = "False"
										ElseIf blnObjDisable = "1" AND UCASE(strData) = "TRUE" Then
												UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Disabled</font>", "Fail"
												Environment.Value("TestStepLog") = "False"
										End If

				End Select
		Else
				UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Link - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If
End Function

'==================================================================================================================================================
' Name of the Function     			  : OperateOnWebRadioGroup
' Description       		   		 	     : This function is used to select or retrieve the values from the WebRadioGroup object
' Date and / or Version       	    : 
' Example Call							 : OperateOnWebRadioGroup("Browser(""TIPS"").Page(""TitleGrid"")","FeelingLucky", "SELECT", "YES")
'==================================================================================================================================================
Function OperateOnWebRadioGroup(strObjectHierarchy, strObjectName, strAction_Name, strData)
Wait(1)
	    Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.WebRadioGroup(strObjectName)
		StepStartTime = Time
		If fnc_wait(ActualObject) = "True" Then
				If instr(strData,"fn_")>0 Then
					strData = Environment.Value(strData)
            	End If
				Select Case strAction_Name
						  Case "SELECT"
							  ObjectHierarchy.WebRadioGroup(strObjectName).Select strData
							  UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Radio button "& Mid(strObjectName,4,Len(strObjectName)-3) & " is selected successfully", "Done"
						  Case "CHECKEXIST"
							  strExp = strData
							  strActual = ObjectHierarchy.WebRadioGroup(strObjectName).Exist
							  If UCASE(strExp) = UCASE(CStr(strActual)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object exists as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strActual)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object does not exist</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  ElseIf UCASE(strExp) = UCASE(CStr(strActual)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object does not exist as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strActual)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object exists</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
							  '****Start***** Added by Suparna on 5 July 2011*******************
							Case "CHECKENABLED"
								blnObjDisable= ObjectHierarchy.WebRadioGroup(strObjectName).GetROProperty("disabled")
                            	If blnObjDisable = "1" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Disabled as expected</font>", "Pass"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Enabled as expected</font>", "Pass"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Enabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								ElseIf blnObjDisable = "1" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Disabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								End If
								  '****End***** Added by Suparna on 5 July 2011*******************
				End Select
		Else
				UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Radio Group - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If
	   
End Function

'==================================================================================================================================================
' Name of the Function     			  : OperateOnWebElement
' Description       		   		 	     : This function is used to select or retrieve the values from the WebElement object
' Date and / or Version       	    : 
' Example Call							 : OperateOnWebElement("Browser(""TIPS"").Page(""TitleGrid"")","GOTO", "CLICK", "")
'==================================================================================================================================================
Function OperateOnWebElement(strObjectHierarchy, strObjectName, strAction_Name, strData)

	    Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.WebElement(strObjectName)
		StepStartTime = Time
		If fnc_wait(ActualObject) = "True" Then
				Select Case strAction_Name
						  Case "CLICK"
							  ObjectHierarchy.WebElement(strObjectName).Click
							  UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Element "& Mid(strObjectName,4,Len(strObjectName)-3) &" is clicked successfully", "Done"
						  Case "CHECKEXIST"
							  strExp = strData
							  strAct = ObjectHierarchy.WebElement(strObjectName).Exist(1)
							  If UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object exists as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object does not exist</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  ElseIf UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object does not exist as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object exists</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
				End Select
		Else
				UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The WebElement - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If
End Function

'==================================================================================================================================================
' Name of the Function     			  : OperateOnImage
' Description       		   		 	     : This function is used to select or retrieve the values from the Image object
' Date and / or Version       	    : 
' Example Call							 : OperateOnImage("Browser(""TIPS"").Page(""TitleGrid"")","GOTO", "CLICK", "")
'==================================================================================================================================================
Function OperateOnImage(strObjectHierarchy, strObjectName, strAction_Name, strData)

	    Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.Image(strObjectName)
		StepStartTime = Time
		If fnc_wait(ActualObject) = "True" Then
				Select Case strAction_Name
						  Case "CLICK"
						  	  Wait(2)
							  ActualObject.Click
							  UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The image " & Mid(strObjectName,4,Len(strObjectName)-3) & " is clicked successfully", "Done"
						  Case "CHECKEXIST"
							  strExp = strData
							  strAct = ObjectHierarchy.Image(strObjectName).Exist
							  If UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object exists as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object does not exist</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  ElseIf UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object does not exist as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object exists</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
				End Select
		Else
				UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Image - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If
End Function

'==================================================================================================================================================
' Name of the Function     			  : OperateOnWebCheckBox
' Description       		   		 	     : This function is used to select or retrieve the values from the WebCheckBox object
' Date and / or Version       	    : 
' Example Call							 : OperateOnWebCheckBox("Browser(""TIPS"").Page(""TitleGrid"")","SELECT", "CHECK", "")
'==================================================================================================================================================
Function OperateOnWebCheckBox(strObjectHierarchy, strObjectName, strAction_Name, strData) 

	    Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.WebCheckBox(strObjectName)
		StepStartTime = Time
		If fnc_wait(ActualObject) = "True" Then
				Select Case strAction_Name
						  Case "CHECK"
								If  ObjectHierarchy.WebCheckBox(strObjectName).GetROProperty("checked") = 0 Then
										ObjectHierarchy.WebCheckBox(strObjectName).Set "ON"
										UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Check box "& Mid(strObjectName,4,Len(strObjectName)-3) &" is checked successfully", "Done"
								Else
										UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Check box "& Mid(strObjectName,4,Len(strObjectName)-3) &" is checked successfully", "Done"
								End If
						  Case "UNCHECK"
								If  ObjectHierarchy.WebCheckBox(strObjectName).GetROProperty("checked") = 1 Then
										ObjectHierarchy.WebCheckBox(strObjectName).Set "OFF"
										UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Check box "& Mid(strObjectName,4,Len(strObjectName)-3)&" is unchecked successfully", "Done"
								 Else
										UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Check box"& Mid(strObjectName,4,Len(strObjectName)-3) &" is unchecked successfully", "Done"
								End If
'***************************** Added by Suresh\ on 28/07/2011 *************************************************
						Case "CHECKEXIST"
							  strExp = strData
							  strAct = ActualObject.Exist
							  If UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object exists as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object does not exist</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  ElseIf UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object does not exist as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object exists</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
						Case "CHECKENABLED"
								blnObjDisable= ActualObject.GetROProperty("disabled")
                            	If blnObjDisable = "1" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Disabled as expected</font>", "Pass"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Enabled as expected</font>", "Pass"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Enabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								ElseIf blnObjDisable = "1" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Disabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								End If
						 Case "CHECKCHECKED"
							 			blnObjDisable= ActualObject.GetROProperty("checked")
                            	If blnObjDisable = "1" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Checked as expected</font>", "Pass"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Not Checked as expected</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Not Checked as expected</font>", "Pass"
                               	ElseIf blnObjDisable = "1" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Checked</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								End If
'**************************************************************************************************************************************
				End Select
		Else
				UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Checkbox - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If
End Function


'==================================================================================================================================================
' Name of the Function     			  : OperateOnWebFile
' Description       		   		 	     : This function is used to enter or retrieve the values on / from the WebFile object
' Date and / or Version       	    : 
' Example Call							 : OperateOnWebFile("Browser(""TIPS"").Page(""TitleGrid"")","LoginName", "SETVALUE", "Test123")
'==================================================================================================================================================
Function OperateOnWebFile(strObjectHierarchy, strObjectName, strAction_Name, strData)
   
		If  strAction_Name <> "STOREVALUE" AND Instr(strData, "VAR_") = 1 Then				' If the value to be taken from already saved variable
			strData = Environment.Value(strData)
		End If

		If  Instr(strData, "fn_") = 1Then
			strData = Environment.Value(strData)
		End If

		Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.WebFile(strObjectName)
		StepStartTime = Time
        If fnc_wait(ActualObject) = "True" Then
				Select Case strAction_Name
					  Case "SETVALUE"
							  ObjectHierarchy.WebFile(strObjectName).Set strData
							  UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Value: <font color=""blue""><b><i>" & strData & "</i></b></font> is entered", "Done"					
					  Case "CHECKEXIST"
							  strExp = strData
							  strActual = ObjectHierarchy.WebFile(strObjectName).Exist
							  If UCASE(strExp) = UCASE(CStr(strActual)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object exists as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strActual)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object does not exist</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  ElseIf UCASE(strExp) = UCASE(CStr(strActual)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object does not exist as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strActual)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object exists</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
					Case "CLICK"
								ObjectHierarchy.WebFile(strObjectName).Click
								UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Webfile is Clicked", "Done"
				End Select
        Else
                UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The WebFile - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If

End Function

'==================================================================================================================================================
' Name of the Function     			  : OperateOnSAPEdit
' Description       		   		 	     : This function is used to enter or retrieve the values on / from the SAPEdit object
' Date and / or Version       	    : 
' Example Call							 : OperateOnSAPEdit("SAPGuiSession("Session").SAPGuiWindow("SAP Easy Access")","LoginName", "SETVALUE", "Test123")
'==================================================================================================================================================
Function OperateOnSAPEdit(strObjectHierarchy, strObjectName, strAction_Name, strData)
   
		If  strAction_Name <> "STOREVALUE" AND Instr(strData, "VAR_") = 1 Then				' If the value to be taken from already saved variable
			strData = Environment.Value(strData)
		End If

		If  Instr(strData, "fn_") = 1Then
			strData = Environment.Value(strData)
		End If

		Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.SAPGuiEdit(strObjectName)
		StepStartTime = Time
		If fnc_wait(ActualObject) = "True" Then
				Select Case strAction_Name
					  Case "SETVALUE"
								ObjectHierarchy.SAPGuiEdit(strObjectName).Set strData
								UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Value: <b><i>" & strData & "</i></b> is enterd", "Done"
					  Case "COMPAREVALUE"
								strExp = strData
								strActual = ObjectHierarchy.SAPGuiEdit(strObjectName).GetROProperty("Text")
								If Trim(UCASE(strExp)) = Trim(UCASE(strActual)) Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Both Expected and Actual values are matching</font>", "Pass"
								Else
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Value mismatch, Actual value is: <i>" & strActual & "</i></font>", "Fail"
									Environment.Value("TestStepLog") = "False"
								End If
				    Case "SETPASSWORD"
								ObjectHierarchy.SAPGuiEdit(strObjectName).SetSecure strData
								UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Password is entered", "Done"
					Case "CHECKENABLED"	   
								blnObjDisable= ObjectHierarchy.SAPGuiEdit(strObjectName).GetROProperty("disabled")
								If blnObjDisable <> strData AND strData = "0" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Disabled as expected</font>", "Pass"
								ElseIf blnObjDisable <> strData AND strData = "1" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Enabled as expected</font>", "Pass"
								ElseIf blnObjDisable = strData AND strData = "0" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Enabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								ElseIf blnObjDisable = strData AND strData = "1" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Disabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								End If
				End Select
		Else
				UpdateReport "TESTSTEP", "", "<font color=""red"">"& Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Textbox - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If

End Function


'==================================================================================================================================================
' Name of the Function     			  : OperateOnSAPButton
' Description       		   		 	     : This function is used to enter or retrieve the values on / from the SAPButton object
' Date and / or Version       	    : 
' Example Call							 : OperateOnSAPButton("SAPGuiSession("Session").SAPGuiWindow("SAP Easy Access")","OK", "CLICK", "")
'==================================================================================================================================================
Function OperateOnSAPButton(strObjectHierarchy, strObjectName, strAction_Name, strData)
   
		Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.SAPGuiButton(strObjectName)

		StepStartTime = Time
		If  fnc_wait(ActualObject) = "True" Then
				Select Case strAction_Name
						Case "CLICK"
								ObjectHierarchy.SAPGuiButton(strObjectName).Click
								UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Button: <b><i>" & strObjectName & "</i></b> is clicked", "Done"
						Case "CHECKENABLED"
								blnObjDisable= ObjectHierarchy.SAPGuiButton(strObjectName).GetROProperty("disabled")
								If blnObjDisable <> strData AND strData = "0" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Button is Disabled as expected</font>", "Pass"
								ElseIf blnObjDisable <> strData AND strData = "1" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Button is Enabled as expected</font>", "Pass"
								ElseIf blnObjDisable = strData AND strData = "0" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Button is Enabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								ElseIf blnObjDisable = strData AND strData = "1" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Button is Disabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								End If
			    End Select
		Else
				UpdateReport "TESTSTEP", "", "<font color=""red"">"& Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Button - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If

End Function

'==================================================================================================================================================
' Name of the Function     			  : OperateOnSAPRadioGroup
' Description       		   		 	     : This function is used to select or retrieve the values from the SAPRadioGroup object
' Date and / or Version       	    : 
' Example Call							 : OperateOnSAPRadioGroup("SAPGuiSession("Session").SAPGuiWindow("SAP Easy Access")","FeelingLucky", "SELECT", "")
'==================================================================================================================================================
Function OperateOnSAPRadioGroup(strObjectHierarchy, strObjectName, strAction_Name, strData)

	    Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.SAPGuiRadioButton(strObjectName)
		StepStartTime = Time
		If  fnc_wait(ActualObject) = "True"  Then
				Select Case strAction_Name
					  Case "SETVALUE"
							  ObjectHierarchy.SAPGuiRadioButton(strObjectName).Set
							  UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Radio button is selected", "Done"
				End Select
	    Else
				UpdateReport "TESTSTEP", "", "<font color=""red"">"& Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Radio button - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If		
	   
End Function

'==================================================================================================================================================
' Name of the Function     			  : OperateOnSAPTabStrip
' Description       		   		 	     : This function is used to select or retrieve the Tab from the SAP Application Screen
' Date and / or Version       	    : 
' Example Call							 : OperateOnSAPTabStrip("SAPGuiSession("Session").SAPGuiWindow("SAP Easy Access")","Tab1", "SELECT", "Tab1")
'==================================================================================================================================================
Function OperateOnSAPTabStrip(strObjectHierarchy, strObjectName, strAction_Name, strData)

	    Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.SAPGuiTabStrip(strObjectName)
		StepStartTime = Time
	     If  fnc_wait(ActualObject) = "True"  Then
				Select Case strAction_Name
						  Case "SELECTVALUE"
							  ObjectHierarchy.SAPGuiTabStrip(strObjectName).Select strData
							  UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Tab is selected", "Done"
				End Select
		Else
			   UpdateReport "TESTSTEP", "", "<font color=""red"">"& Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Tab - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
			   Environment.Value("TestStepLog") = "False"
			   Environment.Value("TestObjectFlag") = "False"
		End If
End Function


'==================================================================================================================================================
' Name of the Function     			  : OperateOnSAPGUITable
' Description       		   		 	     : This function is used to enter the values on  the Cell
' Date and / or Version       	    : 
' Example Call							 : OperateOnSAPGUITable("SAPGuiSession("Session").SAPGuiWindow("SAP Easy Access")","Table", "SETCELLVALUE", "1, Material, P-103")
'==================================================================================================================================================
Function OperateOnSAPGUITable(strObjectHierarchy, strObjectName, strAction_Name, strData)

	    Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.SAPGuiTable(strObjectName)
		StepStartTime = Time
        If  fnc_wait(ActualObject) = "True"  then
					Select Case strAction_Name
							  Case "SETCELLVALUE"
								  Input =Split(strData,",")		
								  int_row_no = int(Input(0))
								  ColName = Input(1)
								  If   Instr(Input(2), "VAR_") = 1 Then				' If the value to be taken from already saved variable
										inputValue = Environment.Value(Input(2))
								  Else
										inputValue = Input(2)
								  End If
								  ActualObject.SetCellData int_row_no,ColName,inputValue
								  Valueinputed = ActualObject.GetCellData (int_row_no,ColName)
								  If  Valueinputed = inputValue Then
											UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The  value " & inputValue & " is enterd into the Table Cell", "Done"
								  Else
											UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Failed to Enter Cell value", "FAIL"
								 End If
					End Select
		Else
		       UpdateReport "TESTSTEP", "", "<font color=""red"">"& Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Table - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
			   Environment.Value("TestStepLog") = "False"
			   Environment.Value("TestObjectFlag") = "False"
		End If			
	   
End Function

'==================================================================================================================================================
' Name of the Function     			  : OperateOnSAPGUIOKCode
' Description       		   		 	     : This function is used  to  enter Transaction code in SAPGuiOKCode Field
' Date and / or Version       	    : 
' Example Call							 : OperateOnSAPGUIOKCode("SAPGuiSession("Session").SAPGuiWindow("SAP Easy Access")","Enter", "SET", "ABC")
'==================================================================================================================================================
Function OperateOnSAPGUIOKCode(strObjectHierarchy, strObjectName, strAction_Name, strData)

	    Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.SAPGuiOKCode(strObjectName)
		StepStartTime = Time
		If  fnc_wait(ActualObject)="True" Then
				Select Case strAction_Name
						  Case "SETVALUE"
							  ActualObject.Set strData
							  UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The  value is entered in field", "Done"
				End Select
		 Else
		       UpdateReport "TESTSTEP", "", "<font color=""red"">"& Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Textbox - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
			   Environment.Value("TestStepLog") = "False"
			   Environment.Value("TestObjectFlag") = "False"
		End If			
	   
End Function

'==================================================================================================================================================
' Name of the Function     			  : ReadSAPStatusbar
' Description       		   		 	     : This function is used  to  Get value from Statubar
' Date and / or Version       	    : 
' Example Call							 : ReadSAPStatusbar("SAPGuiSession("Session").SAPGuiWindow("SAP Easy Access")","SAPStatusbar", "GETVALUE", "")
'==================================================================================================================================================
Function ReadSAPStatusbar(strObjectHierarchy, strObjectName, strAction_Name, strData)
	    Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.SAPGuiStatusBar(strObjectName)
		StepStartTime = Time
		If  fnc_wait(ActualObject) ="True" Then
				Select Case strAction_Name
						  Case "GETVALUE"
							    txtStatus=ActualObject.GetROProperty("text") 
								If  txtStatus <> "" Then
										txtArray = split(txtStatus," ")
										For Each intValue in txtArray 
											If IsNumeric(intValue) then
												   Environment.Value(strData) =  intValue
												   Exit for
											End If
										Next
										If Isnumeric(intValue) AND Instr(txtStatus, "does not exist ") = 0 Then
											   UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") & "</font>" , "<font color=""green"">" & StepStartTime &"</font>" , "<font color=""green"">" & Time&"</font>", "<font color=""green"">The message <b><i>" & txtStatus & "</i></b> appeard on the status bar</font>", "Pass"
										Else
											   UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>" , "<font color=""red"">" & StepStartTime &"</font>" , "<font color=""red"">" & Time&"</font>", "<font color=""red"">The message <b><i>" & txtStatus & "</i></b> appeard on the status bar</font>", "Fail"			
										End If
								ElseIf Instr(txtStatus, "does not exist ") > 0 OR txtStatus = "" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>" , "<font color=""red"">" & StepStartTime &"</font>" , "<font color=""red"">" & Time&"</font>", "<font color=""red"">The message <b><i>" & txtStatus & "</i></b> Appeard</font>", "Fail"			
								End If
				End Select
		Else
			   UpdateReport "TESTSTEP", "", "<font color=""red"">"& Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Statusbar - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
			   Environment.Value("TestStepLog") = "False"
			   Environment.Value("TestObjectFlag") = "False"
		End If		
End Function

'==================================================================================================================================================
' Name of the Function     			  : OperateOnSAPGrid
' Description       		   		 	     : This function is used  to  Get row value from grid
' Date and / or Version       	    : 
' Example Call							 : OperateOnSAPGrid("SAPGuiSession("Session").SAPGuiWindow("SAP Easy Access")","SAPGrid", "SELECTCELL", "")
'==================================================================================================================================================
Function OperateOnSAPGrid(strObjectHierarchy, strObjectName, strAction_Name, strData)
	    Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.SAPGuiGrid(strObjectName)
		StepStartTime = Time
		If fnc_wait(ActualObject)="True" Then
				Select Case strAction_Name
						  Case "SELECTCELL"
							    txtStatus=ActualObject.GetROProperty("text") 
								Input =Split(strData,",")		
								int_row_no = int(Input(0))
								ColName = Input(1)
								ActualObject.ActivateCell int_row_no,ColName
								UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Specified row is selected", "Done"
				End Select
		Else
			    UpdateReport "TESTSTEP", "", "<font color=""red"">"& Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Grid - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
			    Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
	    End If			
End Function

'==================================================================================================================================================
' Name of the Function     			  : OperateOnSAPTree
' Description       		   		 	     : This function is used  to  Get row value from grid
' Date and / or Version       	    : 
' Example Call							 : OperateOnSAPTree("SAPGuiSession("Session").SAPGuiWindow("SAP Easy Access")","SAPTree", "SELECTNODE", "Test123")
'==================================================================================================================================================
Function OperateOnSAPTree(strObjectHierarchy, strObjectName, strAction_Name, strData)
	    Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.SAPGuiTree(strObjectName)
		StepStartTime = Time
		If  fnc_wait(ActualObject)="True" Then
				Select Case strAction_Name
					  Case "SELECTNODE"
							ActualObject.Collapse strData
							ActualObject.ActivateNode strData
							If err.number =  -2147220983 Then
								UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, " Node does not exist", "Fail"
							Else
								Reporter.ReportEvent micDone, "Select  the Node  -  "&strObjectName , "   Node selected in Tree view" 
								UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "  Node selected in Tree view", "Done"
							End If
				End Select
		Else
				UpdateReport "TESTSTEP", "", "<font color=""red"">"& Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Tree Node - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If		
End Function

'==================================================================================================================================================
' Name of the Function                                                    : fn_LauchURL
' Description                                                                                            : This function is used  to  Launch the specified Application URL
' Date and / or Version                       : 
' Example Call                                                                                                     : 
'==================================================================================================================================================
Function fn_LauchURL(strScenarioName, strDataVal)
                                                StepStartTime = Time
                                                strData = strDataVal
'                                               strObjectHierarchy = Datatable.value("APP_SCREEN_NAME", strScenarioName)                              ' Object Hierarchy value ex: Browser("Login").Page("Login")
'                                               strObject = Datatable.Value("OBJECT", strScenarioName)
                                                Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
                                                strData = Split(strDataVal,"\")
                                                SystemUtil.Run "C:\MCS_BRILLIO\"&strData(0),strData(1),,,3
'                                               ActualObject.Sync
                                                UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Application URL is launched successfully", "Done"
End Function

'==================================================================================================================================================
' Name of the Function     			  : fn_GetCellValueFromWebTable
' Description       		   		 	     : This function is used  to get the specified cell values from the Web Table
' Date and / or Version       	    : 
' Example Call							 : 
'==================================================================================================================================================
Function fn_GetCellValueFromWebTable(strScenarioName, strDataVal)
'			Wait(10)
			StepStartTime = Time
			strData = strDataVal
			strObjectHierarchy = Datatable.value("APP_SCREEN_NAME", strScenarioName)		' Object Hierarchy value ex: Browser("Login").Page("Login")
			strObject = Datatable.Value("OBJECT", strScenarioName)
			Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)

			If strObject <> ""  Then
					strActualObject = strObjectHierarchy & "." & strObject
			ElseIf strObject = "" Then
					strActualObject = strObjectHierarchy
			End If

			Set ActualObject = Eval(strActualObject)
			Call fnc_wait(ActualObject)
			strDataSplit = Split(strData, ",")
			strRowNum = strDataSplit(0)
			strColNum = strDataSplit(1)
			strCellData =  ActualObject.GetCellData(strRowNum, strColNum)
            Environment.Value("fn_GetCellValueFromWebTable") = Trim(strCellData)
			UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The record: <b><i>" & Environment.Value("fn_GetCellValueFromWebTable") & "</i></b> is fetched from the Table", "Done"
End Function

'==================================================================================================================================================
' Name of the Function     			  : fn_ClickObjectinWebTable
' Description       		   		 	     : This function is used  to click on the specified (Button or Link or Image) object inside the WebTable
' Date and / or Version       	    : 
' Example Call							 : 
'==================================================================================================================================================
Function fn_ClickObjectinWebTable(strScenarioName, strDataVal)
			StepStartTime = Time
			strData = strDataVal
			strObjectHierarchy = Datatable.value("APP_SCREEN_NAME", strScenarioName)		' Object Hierarchy value ex: Browser("Login").Page("Login")
			strObject = Datatable.Value("OBJECT", strScenarioName)
			Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)

			If strObject <> ""  Then
					strActualObject = strObjectHierarchy & "." & strObject
			ElseIf strObject = "" Then
					strActualObject = strObjectHierarchy
			End If

			Set ActualObject = Eval(strObjectHierarchy)
			Call fnc_wait(ActualObject)

			strDataSplit = Split(strData, ",")
			strRowNum = strDataSplit(0)
			strColNum = strDataSplit(1)
			strObjectName = strDataSplit(2)
			strIndexvalue = strDataSplit(3)

'			If ActualObject.ChildItem(strRowNum, strColNum, strObjectName, strIndexvalue).Exist Then
                    ActualObject.ChildItem(strRowNum, strColNum, strObjectName, strIndexvalue).Click
					UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The item: <b><i>" & strObjectName & "</i></b> present in the cell - (" & strRowNum & ","  & strColNum & ") is clicked", "Done"
'			Else
'					UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Item: <b><i>" & strObjectName & "</i></b> do not exist in the specified cell - (" & strRowNum & ","  & strColNum & ")", "Fail"
'					Environment.Value("TestStepLog") = "False"
'			End If
			
End Function

'==================================================================================================================================================
' Name of the Function     			  : fn_ValidateTextinWebTable
' Description       		   		  : This function is used  to verify the text inside the WebTable cell
' Date and / or Version       	    : 
' Example Call							 : 
'==================================================================================================================================================
Function fn_ValidateTextinWebTable(strScenarioName, strDataVal)
			StepStartTime = Time
			strData = strDataVal
			strObjectHierarchy = Datatable.value("APP_SCREEN_NAME", strScenarioName)		' Object Hierarchy value ex: Browser("Login").Page("Login")
			strObject = Datatable.Value("OBJECT", strScenarioName)
			Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)

			If strObject <> ""  Then
					strActualObject = strObjectHierarchy & "." & strObject
			ElseIf strObject = "" Then
					strActualObject = strObjectHierarchy
			End If

           		If  Instr(strData, "fn_") = 1 Then				' If this value is a return parameter of any function
					strData = Environment.Value(strData)
			End If

			Set ActualObject = Eval(strActualObject)
			Call fnc_wait(ActualObject)
            		intRow = ActualObject.RowCount
            		For intCurRow = 1 to intRow
					intCol = ActualObject.ColumnCount(intCurRow)
					For intCurCol = 1 to intCol
							strCellValue = ActualObject.GetCellData(intCurRow,intCurCol)
							If UCASE(Trim(strCellValue)) =  UCASE(Trim(strData))Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") & "</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The Text: <b><i>""" & strData & """</i></b> appeared in the Table</font>", "Pass"
                                    					Exit Function
								Else
									blnMatchNotFound = "True"
							End If
					Next
			Next
			If blnMatchNotFound = "True" Then
                    			UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Text: <b><i>""" & strData & """</i></b> did not appear in the Table</font>", "Fail"
					Environment.Value("TestStepLog") = "False"
			End If
End Function

'==================================================================================================================================================
' Name of the Function     			  : fn_ReportOnNavigation
' Description       		   		 	     : This function is used  to Report the successful navigation of page.
' Date and / or Version       	    : 
' Example Call							 : 
'==================================================================================================================================================
Function fn_ReportOnNavigation(strScenarioName, strDataVal)
			StepStartTime = Time
			strData = strDataVal
			strObjectHierarchy = Datatable.value("APP_SCREEN_NAME", strScenarioName)		' Object Hierarchy value ex: Browser("Login").Page("Login")
			strObject = Datatable.Value("OBJECT", strScenarioName)
			Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
			strDataSplit = Split(strData, ",")
			strDataOne = strDataSplit(0)
			strDataTwo = strDataSplit(1)
'			wait(2)
			If strObject <> ""  Then
					strActualObject = strObjectHierarchy & "." & strObject
			ElseIf strObject = "" Then
					strActualObject = strObjectHierarchy
			End If
			
			Set ActualObject = Eval(strActualObject)
			Call fnc_wait(ActualObject)

            If  ActualObject.Exist(1) Then
					UpdateReport "TESTSTEP", "", "<font color=""green"">" & strDataOne & "</font>","<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">" & strDataTwo & "</font>", "Pass"
			Else
					UpdateReport "TESTSTEP", "", "<font color=""red"">" & strDataOne& "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Navigation is unsuccessful</font>" , "Fail"
					Environment.Value("TestStepLog") = "False"
					Environment.Value("TestObjectFlag") = "False"
			End If
End Function

'==================================================================================================================================================
' Name of the Function     			  : fn_UpdateValueGlobalParameter
' Description       		   		 : This function is used  to update a parameter into the globla parameter sheet
' Date and / or Version       	    : 25-March-2011
'Author                             : Febin Mathew
' Example Call						: 
'==================================================================================================================================================
Function fn_UpdateValueGlobalParameter(strDataVal)
            strData = strDataVal
			strSplitData = Split(strData,",")
            strColumn = Trim(strSplitData(0))
			strGlobalValue = Trim(strSplitData(1))
			If  Instr(strGlobalValue, "fn_") = 1 Then				' If this value is a return parameter of any function
				strGlobalValue = Environment.Value(strGlobalValue)
			End If
			DataTable.GetSheet("GLOBALPARAMETERS").SetCurrentRow(1)
			DataTable.Value(strColumn,"GLOBALPARAMETERS") = strGlobalValue
			Environment.Value("fn_UpdateValueGlobalParameter") = "True" ' Added by Suresh on 02/05/2012
			'Added by Suparna sheet  is exported 
			'DataTable.ExportSheet Environment.Value("ExcelPath") & "Parameters.xls", "GLOBALPARAMETERS"
			'DataTable.ExportSheet Environment.Value("Resource_Path") & "Parameters.xls", "GLOBALPARAMETERS"      'Updated by sURESH ON 02/05/2012
			 
End Function


'==================================================================================================================================================
' Name of the Function     			  : fn_CloseBrowser
' Description       		   		 	     : This function is used  to Close the Browser
' Date and / or Version       	    : 
' Example Call							 : 
'==================================================================================================================================================
Function fn_CloseBrowser(strScenarioName, strDataVal)

			On Error Resume Next
			StepStartTime = Time
			strBrowserName = Datatable.Value("INPUTDATA_PARAMETER", strScenarioName)
			Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
			If Browser(strBrowserName).Exist(10) Then
				Wait(1)
                Browser(strBrowserName).CloseAllTabs
				UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The browser is closed" , "Done"
			End If
			
End Function

'==================================================================================================================================================
' Name of the Function     			  : fn_CloseWindow
' Description       		   		 	     : This function is used  to Close the specified application window
' Date and / or Version       	    : 
' Example Call							 : 
'==================================================================================================================================================
Function fn_CloseWindow(strScenarioName,strDataVal)
			On Error Resume Next
			StepStartTime = Time
			strObjectHierarchy = Datatable.value("APP_SCREEN_NAME", strScenarioName)		' Object Hierarchy value ex: Browser("Login").Page("Login")
			Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
			Set ActualObject = Eval(strObjectHierarchy)
			Call fnc_wait(ActualObject)
			ActualObject.Close()
			If Err.Number = 0 Then
				UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Application Window is Closed" , "Done"
			ElseIf Err.Number <> 0 Then
				Call fn_CloseAllPBApplications(strScenarioName,strDataVal)
				UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Application Window is Closed" , "Done"
			End If	
End Function

'==================================================================================================================================================
' Name of the Function     			  : fn_EnterKey
' Description       		   		 	     : This function is used  press 'Enter' key
' Date and / or Version       	    : 
' Example Call							 : 
'==================================================================================================================================================
Function fn_EnterKey(strScenarioName,strDataVal)
			Wait(1)
			strObjectHierarchy = Datatable.value("APP_SCREEN_NAME", strScenarioName)
			strObject = Datatable.Value("OBJECT", strScenarioName)
			strActualObject = strObjectHierarchy & "." & strObject
			Set ActualObject = Eval(strActualObject)
			ActualObject.Click
			
			Set WshShell = CreateObject("WScript.Shell")
			WshShell.Sendkeys "{ENTER}"
End Function


'==================================================================================================================================================
' Name of the Function     			  : fn_LaunchApp
' Description       		   		 	     : This function is used  to launch the SAP Application
' Date and / or Version       	    : 
' Example Call							 : 
'==================================================================================================================================================
Function fn_LaunchApp(strScenaraioName,strDataVal)

			SystemUtil.Run strDataVal,"","","Open"
			Wait(2)

End Function

'==================================================================================================================================================
' Name of the Function     			  : fn_CloseAllBrowser
' Description       		   		 	     : This function is used  to Close the Browser
' Date and / or Version       	    : 
' Example Call							 : 
'==================================================================================================================================================
Function fn_CloseAllBrowser()
		SystemUtil.CloseDescendentProcesses
End Function

'==================================================================================================================================================
' Name of the Function     			  : fn_FindLocalParamRowNumber
' Description       		   		  : This function is used  to find the row number of the TC in LocalParameter sheet
' Date and / or Version       	   	 : 
' Example Call						 : 
'==================================================================================================================================================
Function fn_FindLocalParamRowNumber(strTestCaseName)
		DataTable.GetSheet("LOCALPARAMETERS").SetCurrentRow(1)
	 	For intCounter = 1 to DataTable.GetSheet("LOCALPARAMETERS").GetRowCount
				If  DataTable.Value("TESTCASE_NAME", "LOCALPARAMETERS") = strTestCaseName Then
						Environment.Value("LocalParamRow") = intCounter
						Exit For
				Else
						DataTable.GetSheet("LOCALPARAMETERS").SetNextRow
				End If
		Next
End Function 

'==================================================================================================================================================
' Name of the Function     			  : UploadResultFile
' Description       		   		 	     : This function is used  to Upload the result files to QC after the execution
' Date and / or Version       	    : 
' Example Call							 : 
'==================================================================================================================================================
Function UploadResultFile()
		Dim o_CurrentRun, o_AttachmentsFact, o_Att, o_ExtStr
		' Zip the result HTMLfile
		zipFile = Environment.Value("ResultPath") & Environment.Value("HTMLFoldername") & ".Zip"
	    Set fso = CreateObject("Scripting.FileSystemObject")
		Set ts = fso.OpenTextFile(zipFile, 2,vbtrue)
		BlankZip = "PK" & Chr(5) & Chr(6)
		For x = 0 to 17
		BlankZip = BlankZip & Chr(0)
		Next
		ts.Write BlankZip
		Set ts = Nothing
        Set sa = CreateObject("Shell.Application")
		Set zip= sa.NameSpace(zipFile)
		Set Fol=sa.NameSpace(Environment.Value("ResultPath") & Environment.Value("HTMLFoldername"))
		zip.CopyHere(Fol)
		zip.CopyHere(Environment.Value("ResultPath") & Environment.Value("strVBSFileName"))
		Wait(5)

		' Upload the result html .zip file to QC Test Instance as an attachment
		v_Path = Environment.Value("ResultPath")
		v_Filename = Environment.Value("HTMLFoldername") & ".Zip"

        Set o_CurrentRun = QCUtil.CurrentRun
        If (o_CurrentRun Is Nothing) Then		' Check that we are running this test from QC, otherwise we can exit
			Exit Function
		End If
	
		If (Right(v_Path,1) = "\") Then v_Path = Left(v_Path, Len(v_Path)-1)		'If the v_Path has a trailing \, remove it.
		
		'now attach the file to the current test
		Set o_AttachmentsFact = o_CurrentRun.Attachments
		Set o_att = o_AttachmentsFact.AddItem(v_Filename)
		o_att.Post
	
		Set o_ExtStr = o_att.AttachmentStorage
	
		o_ExtStr.ClientPath = v_Path
		o_ExtStr.Save v_Filename, true
	
		f_UploadQCAttachment = o_att.Post

End Function


'=====================================================================================================================================================================================
' Name of the Function     			  : fn_CreateViewReportFile
' Description       		   		 	     	: This function is used to create the Unzip Function file.
' Date and / or Version       	      : 27-Sept-2012
' Author										    : Manish
' Input Parameters					     : None
' Example Call							    : Call fn_CreateViewReportFile()
'======================================================================================================================================================================================
Public Function fn_CreateViewReportFile()
	Const ForReading=1
	Const ForWriting=2
	strFilename = "UnZipfunction.vbs"
	If instr( Environment.Value("Resource_Path") , "[QualityCenter]") > 0  Then
			strQCFilePath = Environment.Value("Resource_Path")
			strScriptPath = Environment.Value("ResultPath")
			strUploadFolder = strScriptPath & "DocumentToUpload"
			Environment.Value("strFolderPath") = strUploadFolder
			strUploadFilePath = strUploadFolder & "\" & strFilename ' QC path
			Set objFSO = CreateObject("Scripting.FileSystemObject")
			If NOT objFSO.FolderExists(strUploadFolder) Then
					objFSO.CreateFolder(strUploadFolder)
			End If
			If NOT objFSO.FileExists(strUploadFilePath) Then
					Set qtApp = CreateObject("QuickTest.Application")
					qtApp.Folders.Add(strQCFilePath)
					strTempLoc = PathFinder.Locate(strFilename)

					If Trim(strTempLoc) <> "" Then
							CreateObject("Scripting.FileSystemObject").Movefile  strTempLoc, strUploadFilePath
					End If
					Set qtApp = Nothing				
			End If
			Set objFSO = Nothing	
			strFilePath = strUploadFilePath
	Else
			strScriptPath = Environment.Value("TestDir")
			strIndividualFolder = Split(strScriptPath,"\",-1,1)
			strIntPath =""		
			'Getting Framework Path
			For intCounter = 0 to UBound(strIndividualFolder) - 1
			   strIntPath = strIntPath & strIndividualFolder(intCounter)  & "\"
			   strIntPath = Trim(strIntPath)
			Next
			'For LOacal Execution
			strFilePathPath = strIntPath & "DocumentsToUpload\" 
			strFilePath = strFilePathPath & strFilename	
	End If

	strSourceFilePath = strFilePath
	Environment.Value("strVBSFileName") = "ClickToOpenResultFile.vbs"
	strDestFilePath = Environment.Value("ResultPath") & Environment.Value("strVBSFileName")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set FH_SourceFile = objFSO.OpenTextFile(strSourceFilePath, ForReading, True)
	Set FH_DestFile =  objFSO.OpenTextFile(strDestFilePath, ForWriting, True)
	Do While Not FH_SourceFile.AtEndofStream
			strExpLine  = FH_SourceFile.ReadLine
			 If InStr(strExpLine, "strZipFileName =") > 0 Then
					  strExpLine = "strZipFileName = """& Environment.Value("HTMLFoldername") & """"
			 End If
			 FH_DestFile.WriteLine strExpLine
	Loop

	FH_SourceFile.Close
	FH_DestFile.Close

	Set objFSO = Nothing
	Set FH_SourceFile = Nothing
	Set FH_DestFile = Nothing
End Function


'=====================================================================================================================================================================================
' Name of the Function     			  : fn_UploadParametersSheet
' Description       		   		 	     	: This function is used to upload the Parameters sheet back to QC at the end of execution
' Date and / or Version       	      : 25-Mar-2011
' Author										    : Shrinidhi Holla
' Input Parameters					     : None
' Example Call							    : Call fn_UploadParametersSheet()
'======================================================================================================================================================================================
Function fn_UploadParametersSheet()
		DataTable.DeleteSheet("Global")
		DataTable.DeleteSheet("Action1")
		DataTable.DeleteSheet(Environment.Value("strTestCase"))
		DataTable.DeleteSheet("LOCALPARAMETERS")
		DataTable.Export Environment.Value("ResultPath") & "Parameters.xls"
		strQCResourcePath = Replace(Environment.Value("Resource_Path"), "[QualityCenter]", "")
		Set QCConnection = QCUtil.QCConnection
		Set TreeManager = QCConnection.TreeManager
		Set Node = TreeManager.nodebypath(Trim(strQCResourcePath))
		Set Att = Node.Attachments
		intAttCount = Att.NewList("").Count
		For intCounter = 1 to intAttCount
				strAttName = Att.NewList("").Item(intCounter).Name
				If Instr(strAttName, "_Parameters.xls") > 0 Then
						strAttID = Att.NewList("").Item(intCounter).ID
						Exit For
				End If
		Next
		Att.RemoveItem(strAttID)
		Wait(2)
		Set Attchmt = Att.AddItem(Null)
		Attchmt.FileName = Environment.Value("ResultPath") & "Parameters.xls"
		Attchmt.Type = 1
		Attchmt.Post()    

		Set Att = Nothing
		Set Node = Nothing
		Set TreeManager = Nothing
		Set QCConnection = Nothing
End Function


'=====================================================================================================================================================================================
' Name of the Function     			  : fn_GenerateDateTimeStamp
' Description       		   		 	     	: This function is used to generate the date time stamp and concatinate it with LParam value
' Date and / or Version       	      : 17-May-2011
' Author										    : Shrinidhi Holla
' Input Parameters					     : None
' Example Call							    : Call fn_GenerateDateTimeStamp("testdesc")
'======================================================================================================================================================================================
Function fn_GenerateDateTimeStamp(strDataVal)
	 	strDateTime = Date &"_" &Time
		Set regEx = New RegExp
		regEx.Global = True
		regEx.Pattern = "[/\ \:]"
		strTimeStamp = regEx.replace(strDateTime, "_")
		If InStr(strDataVal,"fn_") > 0 OR InStr(strDataVal,"VAR_") > 0 Then
				strDataVal = Environment.Value(strDataVal)
		End If
		strData =strDataVal & strTimeStamp
		Environment.Value("fn_GenerateDateTimeStamp") = strData
End Function
'*************************************Added by Suresh on 01/08/2011 ***********************************
'==================================================================================================================================================
' Name of the Function     			  : OperateOnWebTable
' Description       		   		 	     : This function is used to click on the Link
' Date and / or Version       	    : 
' Example Call							 : OperateOnWebTable("Browser(""TIPS"").Page(""TitleGrid"").WebTable(""wtbInbox"")","GoTo", "CLICK", "")
'==================================================================================================================================================
Function OperateOnWebTable(strObjectHier, strAction_Name, strData)
		StepStartTime = Time
	    Set ObjectHierarchy = Eval(strObjectHier)
		Call fnc_wait(ObjectHierarchy)
        
		If fnc_wait(ObjectHierarchy) = "True" Then
				Select Case strAction_Name
					Case "VerifyTableWithColumnNames"
							If  strData<>"" Then		
									ArrColumnName=split(strData,",")
									intRowCount = ObjectHierarchy.RowCount
									IntRow =1
									For intRowNo = 1 to intRowCount - 1
											IntColumnCount = ObjectHierarchy.ColumnCount(intRowNo)
											If IntColumnCount > 1 Then
													IntRow= intRowNo
													Exit For
											End If
									Next
									'IntColumnCount= ObjectHierarchy.ColumnCount(1) 
									'IntRow=1   
									'If IntColumnCount=1 Then
									'		IntColumnCount= ObjectHierarchy.ColumnCount(2)        
									'		IntRow=2
									'End If
									For intColName=0 to Ubound(ArrColumnName)
											For intCol=1 to IntColumnCount
													StrCellData=ObjectHierarchy.GetCellData(IntRow,intCol)
												If Instr(UCASE(StrCellData),"ERROR") <> 1 Then													'Added this to handle if cell value is displayed with ERROR value by Basavaraj on 8 March, 2012
													If Instr(StrCellData,ArrColumnName(intColName))>0 Then
															IntFound=0
															Exit For
													Else
															IntFound=1
													End If
												 End If
											Next
											If IntFound=1 Then
													Exit For
											End If
									Next
									If IntFound=0 Then
											UpdateReport "TESTSTEP", "","<font color=""green"">"& Environment.Value("strDescription") & "</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", " <font color=""green""> Column(s)  "&strData&"  are displayed in the table</font> ", "Pass"         
									Else
											UpdateReport "TESTSTEP", "","<font color=""red"">"& Environment.Value("strDescription")& "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", " <font color=""red""> Column(s)  "&strData&"  are not displayed in the table</font> ", "Fail"         
											Environment.Value("TestStepLog") = "False"
											Environment.Value("TestObjectFlag") = "False"
									End If
							Else
									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"),StepStartTime,Time,"No data available for column search", "Done"    
							End if 

					Case "VerifyTableisLoadedwithValues"                								
							arrActualObject=Split(strObjectHier,".")
							strName=arrActualObject(2)
							arrStrName=Split(strName,"(")
							ActualObjName=Replace(arrStrName(1),")","")
							intRowCount=ObjectHierarchy.RowCount
							If intRowCount>1 Then
						       	UpdateReport "TESTSTEP", "", Environment.Value("strDescription"),StepStartTime,Time,"Table: "&ActualObjName&" is having rows of data", "Pass"  
							Else
								Environment.Value("TestStepLog") = "False"
								Environment.Value("TestStepLog") = "False"
								Environment.Value("TestObjectFlag") = "False"
							End If

					Case "CheckRowValueInTable"
							arrstrData=split(strData,",")
							strColName=trim(arrstrData(0))
							strSearchText = trim(arrstrData(1))
                            If InStr(strSearchText,"fn_") = 1 OR InStr(strSearchText,"VAR_") = 1 Then
									strSearchText = Environment.Value(strSearchText)
							End If		
'***************************************Modified by Febin on 12/01/2011 for dynamically changing columns*******************************************
							intRowCount = ObjectHierarchy.RowCount
							For intRow = 1 to intRowCount - 1
									intColumnCount = ObjectHierarchy.ColumnCount(intRow)
									If intColumnCount > 1 Then
											intRCount= intRow
											Exit For
									End If
							Next
'							IntRCount =1   
'							If intColumnCount =1 Then
'								intColumnCount = ObjectHierarchy.ColumnCount(2)        
'								IntRCount =2
'							End If                    
'*************************************End of Modification*******************************************************************************************************    	    
							' To know  the Place of Column in the Table                                         
							For colCount=1 to  intColumnCount
									strCellData=ObjectHierarchy.GetCellData(intRCount,colCount)
                                    If strCellData <> "" AND Instr(UCASE(Trim(strCellData)),UCASE(Trim(strColName))) > 0 Then 
                                    		strStatus= True
											intRow = ObjectHierarchy.RowCount
											For rowCount = intRCount + 1 to intRow
'*********************************************************************Following If added by Febin on 12/02/2011 for not comparing rows having only 1 column****************************************
													If  ObjectHierarchy.ColumnCount(rowCount) > 1Then													
															strActualText = ObjectHierarchy.GetCellData(rowCount,colCount)
															If Instr(strActualText, strSearchText) >= 1 Then
																	Environment.Value("fn_CheckColumnValue") = strActualText
																	UpdateReport "TESTSTEP", "","<font color=""green"">"& Environment.Value("strDescription") & "</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", " <font color=""green"">The value: <b><i>" & strSearchText & "</i></b> is available in the Table.</font>" , "Pass"         
																	strFlag=True
																	Exit For
															Else
																	strFlag=False
															End If
													End If
											Next
													Exit For
											Else
											strStatus= False
									End If							  
							Next           
							
							If strFlag=False Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The value: <b><i>" & strSearchText & "</i></b> is not available in the Table</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
									Environment.Value("TestObjectFlag") = "False"
							End If		  
							If strStatus= False Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Column  : <b><i>" & strColName & "</i></b> is not available in the Table</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
									Environment.Value("TestObjectFlag") = "False"
							End If

				Case "ClickOnaRowinTable"
							strData1 = split(strData,",")
							strColumn = trim(strData1(0))
							If  Instr(strData1(1), "fn_") >= 1 Then                                                           
									strSearchText = Environment.Value(strData1(1))
							Else
									strSearchText=trim(strData1(1))
							End If	
							If  Instr(strSearchText, "VAR_") >= 1 Then                                                           
									strSearchText = Environment.Value(strSearchText)
							End If	
							intColumnCount= ObjectHierarchy.ColumnCount(1)		
							
							intRowCount=  ObjectHierarchy.RowCount			
	'**********Updated on 23 aug 11 by Suparna
							intRCount=1   
							For i = intRCount+1  to intRowCount
								If intColumnCount=1 Then
									intColumnCount= ObjectHierarchy.ColumnCount(i)        
									intRCount=i
								Else
								   Exit For
								End If
							Next
'							If intColumnCount=1 Then
'							intColumnCount= ObjectHierarchy.ColumnCount(2)        
'							intRCount=2
'							End If
	'******************Update End		*****************
							
								   
							'To know  the Place of Column in the Table                                         
							For intcolCount=1 to  intColumnCount
									strCellData=ObjectHierarchy.GetCellData(intRCount,intcolCount)
									If Instr(UCASE(Trim(strCellData)),UCASE(Trim(strColumn))) > 0 Then 
											isColFound=True
											For intRCount=intRCount to intRowCount 
													strActualText=ObjectHierarchy.GetCellData(intRCount,intcolCount)
													If Instr(strActualText, strSearchText) >= 1 Then
															ObjectHierarchy.ChildItem(intRCount,intcolCount,"Link",0).click	
															UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The link <b><i>" & strActualText & "</i></b> is clicked", "Done"						
															strFlag=True
															Exit For
													Else
															strFlag=False
													End If
											Next
											Exit For
									Else
											isColFound=False
									End If		                          
							Next        
						
							If strFlag=False Then		  
									UpdateReport "TESTSTEP", "","<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>","<font color=""red"">The value: <b><i>" & strSearchText & "</i></b> is not available in the Table</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
									Environment.Value("TestObjectFlag") = "False"
							End If	  
							If isColFound=False Then					 
									UpdateReport "TESTSTEP", "","<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>","<font color=""red"">The Column  : <b><i>" & strColumn & "</i></b> is not available in the Table</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
									Environment.Value("TestObjectFlag") = "False"
							End If
					Case "VerifyDataNotPresentinTable"
							strSearchText=Trim(strData)
        					If Instr(strSearchText,"fn_") > 0 OR Instr(strSearchText,"VAR_") >0Then
									strSearchText=Environment.Value(strSearchText)
							Else
									strSearchText= strSearchText
							End If  						
							intRowNumber = ObjectHierarchy.GetRowWithCellText(strSearchText)
							If intRowNumber < 1 Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green""> The row with value:  <b> " & strSearchText & " </b>  not found in the table as expected</font>", "Pass"					
							else
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The row with value: <b> " & strSearchText & " </b>  is present in the table</font>", "Fail"
									Environment.Value("TestStepLog") = "False"	
									Environment.Value("TestObjectFlag") = "False"							
							End If

					Case "VerifyDesiredDataPresentinTable"
							strData = Trim(strData)
							If Instr(strData,"fn_") > 0 OR Instr(strData,"VAR_") >0 Then
									strSearchText = Environment.Value(strData)
							Else
									strSearchText = strData
							End If							
									intRowNumber = ActualObject.GetRowWithCellText(strSearchText)							
							If intRowNumber <> -1 Then
										UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The desired data <b>"& strData &"</b> exists in the table", "Pass"				
							Else	
										UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The desired data <b>"& strData &"</b> does not exists in the table", "Fail"
										Environment.Value("TestStepLog") = "False"
							End If

					Case "VerifyResultTableWithExpectedRowValues"
								arrNumInputData = split(strData,";",-1,1) ' Getting the Count of Column to Verify
								For intNumCol = 0 to  UBound(arrNumInputData)                            
												arrInputData=Split(arrNumInputData(intNumCol),",",-1,1) ' Spliting the Column and ColumnValue
												strColName = arrInputData(0)
												If  Instr(arrInputData(1), "fn_") >= 1 OR Instr(arrInputData(1), "VAR_") >= 1Then                                                            ' If this value is a return parameter of any function
														strExpectedColValue = Environment.Value(arrInputData(1))
												Else
														strExpectedColValue=arrInputData(1)
												End If
'                                                                strExpectedColValue =arrInputData(1)      
'******************************************************************Updated by Febin on 12/07/2011 for dynamically changing columns*****************************************  
												intRowCount = ObjectHierarchy.RowCount
												For intRow = 1 to intRowCount - 1
														intColCount = ObjectHierarchy.ColumnCount(intRow)
														If intColCount > 1 Then
																intRcount= intRow
																Exit For
														End If
												Next
'												intColCount = ObjectHierarchy.ColumnCount(1)
'												intRcount = 1
'												If intColCount = 1 Then
'														intColCount = ObjectHierarchy.ColumnCount(2)
'														intRcount = 2
'												End If 
'*****************************************************************End of Update****************************************************************************************************                                                

												' To know  the Place of Column in the Table                                         
											For intCnt=1 to  intColCount
																strCellData=ObjectHierarchy.GetCellData(intRcount,intCnt)																												
																		If Instr(strCellData,strColName)>0  Then                             
																						intColNum = intCnt   ' Getting the Column Number
																						isColFound = True
																						Exit For
																		Else
																						isColFound=False            
																		End If

												Next
												If isColFound = True Then
																'To Verify the Expected and Actual Column Value
																intRowCount = ObjectHierarchy.RowCount
																If intRowCount > 1 Then		
																				For intRow = intRcount + 1 to intRowCount
																								If  ObjectHierarchy.ColumnCount(intRow) > 1Then																								
																										strActualColValue=ObjectHierarchy.GetCellData(intRow,intColNum)
																										If Instr(UCASE(strActualColValue),"ERROR") <> 1 Then	
																												If  instr(Ucase(Trim(strActualColValue)) ,Ucase(trim(strExpectedColValue))) >=1 Then                                                                                
																																isColMatched = True																																  
																												Else
																																isColMatched = False
																																Exit For
																												End If
																										End If
																								End If
																				Next
																Else
																				UpdateReport "TESTSTEP", "","<font color=""red"">"& Environment.Value("strDescription")& "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", " <font color=""red""> No Rows found  </font>", "Fail"         
																				Environment.Value("TestStepLog") = "False"
																				Environment.Value("TestObjectFlag") = "False"   
																End IF
												Else        
																'Column is Not present in the table
																UpdateReport "TESTSTEP", "","<font color=""red"">"& Environment.Value("strDescription")& "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", " <font color=""red"">" & strColName & "  Column Not Found</font>", "Fail"         
																Environment.Value("TestStepLog") = "False"   
																Environment.Value("TestObjectFlag") = "False"  
												End If    
												If  isColMatched = True Then                                      
																UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,   strColName &" Column  is displayed with Value  "& strExpectedColValue, "Pass"         
												Else                        
																UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime, Time,  strColName &" Column  is displaying  "& strActualColValue  &" instead of  "& strExpectedColValue, "Fail"         
																Environment.Value("TestStepLog") = "False"
																Environment.Value("TestObjectFlag") = "False"   
																Exit For
												End If    
								Next      

 '************************* Added by Basavaraj on Dec 29, 2011**********************************  This function is to verify only the expected value and come out from the loop
					Case "VerifySearchResultDisplayed"
								StepStartTime = Time							
								strDataValue=Split(strData,",")        ' Spliting the Column name and Column Values
								strColumnName = strDataValue(0) ''Column Name
								If instr(strDataValue(1),"fn_")>=1  Then
										strValue = Environment.Value(strDataValue(1))
								Elseif instr(strDataValue(1),"VAR_")>=1  Then
										strValue = Environment.Value(strDataValue(1))
								Else
										strValue = strDataValue(1)
								End If
								'Get the column number
								strRowCount=ObjectHierarchy.RowCount
								strColCount=ObjectHierarchy.ColumnCount(1)
								IntRow=1   
								If strColCount=1 Then
									strColCount= ObjectHierarchy.ColumnCount(2)        
									IntRow=2
											If strColCount=1 Then
											strColCount= ObjectHierarchy.ColumnCount(3)        
											IntRow=3
											End If
								End If
							
								If strRowCount>2 Then		
									For j=1 to strColCount
										If Instr(Ucase(Replace(ObjectHierarchy.GetCellData(IntRow,j)," ","")),Ucase(replace(strColumnName," ","")))>0 Then
											strCellValue=ObjectHierarchy.GetCellData(IntRow+1,j)
											If Instr(Ucase(Trim(strCellValue)),Ucase(Trim( strValue)))>0 Then
												UpdateReport "TESTSTEP", "","<font color=""green"">"& Environment.Value("strDescription")& "</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>","<font color=""green""> The <b> "&strValue&" </b> value exists in the <b>" & strColumnName& " </b> Column</font>", "Pass"  
												strFound=1
												Exit For
											Else
												UpdateReport "TESTSTEP", "","<font color=""red"">"& Environment.Value("strDescription")& "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>","<font color=""red"">The <b> "&strValue&" </b> value does not exists in the <b>" & strColumnName& " </b> Column</font>", "Fail"  
												Environment.Value("TestStepLog") = "False"
												strFound=1
												Exit For
											End If
										Else
											strFound=0
										End If  								
									Next
									If strFound=0 Then
										UpdateReport "TESTSTEP", "","<font color=""red"">"& Environment.Value("strDescription")& "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>","<font color=""red"">" & strColumnName& " does not exists in the search result table</font>", "Fail"  			     
										Environment.Value("TestStepLog") = "False"
									End If
								Else
									UpdateReport "TESTSTEP", "","<font color=""red"">"& Environment.Value("strDescription")& "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>","<font color=""red""> No records found in the search result table</font>", "Fail"          
									Environment.Value("TestStepLog") = "False"
								End If 

'************************************************ Added by Basavaraj *on 29 Dec ********************************************************

					Case "GetDependentCellValue"
							If InStr(strData,";") >= 1 Then '  Since Customer Name Contains the "," , So, to handle that changed the delimitor 
									strData=Split(strData,";")		
							Else
									strData=Split(strData,",")	
							End If
					
							strInpudata = ""
							For iteration = 0 to UBOUND(strData) -1
					'*****************Updated by Febin on 12/07/2011 to take care of Environment Variables**********************************************
									If InStr(Trim(strData(iteration)),"fn_") > 0 OR InStr(Trim(strData(iteration)),"VAR_") > 0 Then											
											strData(iteration) = Environment.Value(Trim(strData(iteration)))
									End If
					'*******************************End of update*****************************************************************************************************
									If  iteration = 0 Then
											strInpudata =strData(iteration) 
									Else
											strInpudata =strInpudata & " , " & strData(iteration) 
									End If
							Next
					
							'To get the Row which contain the Column Name
			'*******************************************Updated by Febin on 12/07/2011 for dynamic columns***********************************************
							intRowCount = ObjectHierarchy.RowCount
							For intRow = 1 to intRowCount - 1
									intColumnCount = ObjectHierarchy.ColumnCount(intRow)
									If intColumnCount > 1 Then
											intRCount= intRow
											Exit For
									End If
							Next
			'				intColumnCount1 = ActualObject.ColumnCount(1)
			'				intColumnCount2 = ActualObject.ColumnCount(2)        
			'				If intColumnCount1=intColumnCount2 Then
			'						intColumnCount = intColumnCount1
			'						IntRCount =1   
			'				Else
			'						intColumnCount = intColumnCount2  
			'						IntRCount =2
			'				End If 
			'********************************************************End of update*********************************************************************************
							' To get the count of columns  need to verify
							NumofCol = UBOUND(strData) /2   'To get  the count of total Number of Columns passed
							ReDim arrColumnNum(NumofCol) '  To collect the Position of Column
							For intColName = 0 to NumofCol  ' Iterate to collect the position of Column
									isColumnFound = False
									For intColIteration = 1 to intColumnCount  ' Iterate to Know the Position of Column
										strColName =ObjectHierarchy.GetCellData(intRCount,intColIteration) '  collecting the Position of Column
									'msgbox strColName &"   Pssed " & strData(intColName*2)
									If  Trim(strData(intColName*2))=""Then
												If  Ucase(Trim(strColName)) = Ucase(Trim(strData(intColName*2))) then
																arrColumnNum(intColName) =  intColIteration ' To get Column Number for all the Column
													isColumnFound = True
													Exit For
												End If
									ElseIf Instr(Ucase(Trim(strColName)),Ucase(Trim(strData(intColName*2)))) > 0 Then														
													arrColumnNum(intColName) =  intColIteration ' To get Column Number for all the Column
													isColumnFound = True
													Exit For
									End If				
									Next
									If isColumnFound = False Then '  To Report Which Column is Not Found in Application.
											UpdateReport "TESTSTEP", "", Environment.Value("strDescription") , StepStartTime ,  Time ,  "The column - "& strData(intColName*2) & " is not found" , "Fail"
											Environment.Value("TestObjectFlag") = "False"
									End If
							Next
							If  isColumnFound = True Then ' If all column found then only proceed		
								' To Get the total Row Count
								intRowCount =  ObjectHierarchy.RowCount
								If  intRowCount >= 1Then '  To Handle if  table is not having any rows		
										For  intRowIteration = IntRCount +1   to intRowCount '  Iterate through each row  to find row which contains Required column Value
												 isColValueFound = False '  To Report  if No such Row found  which sutisty the given condition
												ReDim arrAppColValue(NumOfCol-1)  ' To Collect the Application Column Value ' Taken  "NumOfCol - 1" because dependent number  of column would be one less than total number of column passed
												For intCollIteration = 0 to NumOfCol-1 '  Iterate to match the expected value to AppColValue
														arrAppColValue(intCollIteration) = ObjectHierarchy.GetCellData(intRowIteration,arrColumnNum(intCollIteration) ) ' Storing the AppcolValues for the corresponding columns
														
														If  Instr( Ucase(Trim(arrAppColValue(intCollIteration))),Ucase(Trim(strData(intCollIteration*2 +1)))) > 0 Then  '  Matching the expected value to AppColValue, "intCollIteration*2 +1" will give always Odd number,and ColumnValue  position in Passed strdata would be odd number. 
																isValueFound = True  '  To Know  reqired one criteria is satisfied
														Else
																isValueFound = False
																Exit For
														End If
												Next			
												If  isValueFound = True Then '  if it is true that means required row number is found
														strExpectedColValue = ObjectHierarchy.GetCellData(intRowIteration,arrColumnNum(NumOfCol) )
														Environment.Value("VAR_GetDependentCellValue") = strExpectedColValue
														strExpectedField = strData(UBOUND(strData))
														set regEx = New RegExp
																											regEx.global = true
																											regEx.pattern = "[-?_,.()/:#]"
																											strExpectedField = regEx.replace(strExpectedField, "")
																											strExpectedField = Trim(Replace(strExpectedField, " ", ""))
														Environment.Value("VAR_GetDependentCellValue"&strExpectedField) = strExpectedColValue
														isColValueFound  = True
														UpdateReport "TESTSTEP", "", Environment.Value("strDescription") , StepStartTime ,  Time , strData(UBOUND(strData)) &" with  " & strInpudata  &" is " & strExpectedColValue & " Found and fetched " , "Pass"
														Exit For
												End If  
										Next
										If isColValueFound = False Then 'To Report that  required row  is  not found for given criteria
												UpdateReport "TESTSTEP", "", Environment.Value("strDescription") , StepStartTime ,  Time , "Expected Row value not found for given criteria" , "Fail"
												Environment.Value("TestStepLog") = "False"
												Environment.Value("TestObjectFlag") = "False"
										End If
								Else
											UpdateReport "TESTSTEP", "", Environment.Value("strDescription") , StepStartTime ,  Time , "No Rows Found in the table" , "Fail"
											Environment.Value("TestStepLog") = "False"
											Environment.Value("TestObjectFlag") = "False"
								End If
							End If'

                    Case "GetDoubleDependentCellValue"
									strData = split(strData,",")
									strFirstColName = strData(0)
									strSearchText = strData(1)
									strSecColName=strData(2)
									strExpectedText=strData(3)
									strThirdColName = strData(4)
									intColumnCount= ObjectHierarchy.ColumnCount(1)
									intRowCount=  ObjectHierarchy.rowcount       
'								    To know  the Place of Column in the Table                                         
									For intCnt=1 to  intColumnCount
												strCellData=ObjectHierarchy.GetCellData(1,intCnt)
												If Instr(Trim(strFirstColName),Trim(strCellData))>0 Then 
														strColStatus=True
														Exit For
												Else
														strColStatus=False
												End If
									Next
									If strColStatus=True Then
											For i=1 to intRowCount
													strActualText=ObjectHierarchy.GetCellData(i,j)
													If Instr(strActualText, strSearchText) >= 1 Then
															intRowNumber=i
															strFound=True
															Exit For
													Else
															strFound=False
													End If
											Next
											If strFound=True Then
													For j=1 to  intColumnCount
															strCellData=ObjectHierarchy.GetCellData(1,j)
															If instr(trim(strSecColName),Trim(strCellData))=1 Then 
																		strCol2Status=True
																		intCol2Num=j
																		Exit For
															Else
																		strCol2Status=False
															End If
													Next
													If  strCol2Status=True Then
																For m=1 to intRowCount 
																		strActualText=ObjectHierarchy.GetCellData(m,intCol2Num)
'																		strActualText=ActualObject.GetCellData(intRowNumber,intCol2Num)
																		If Instr(strActualText, strExpectedText) = 1 Then
																				strExpFound=True		
																				intExpRow=m
																				Exit For
																		Else
																				strExpFound=False
																		End If
																Next
																If strExpFound=True Then
																			For l=1 to  intColumnCount
																					strCellData=ObjectHierarchy.GetCellData(1,l)
																					If instr(trim(strThirdColName),Trim(strCellData))=1 Then 
																							strCol3Status=True
																							intLinkCol=l
																							Exit For
																					Else
																							strCo3lStatus=False
																					End If
																			Next
																			If strCol3Status=True then
																				  intExpectedLoanNumber=ObjectHierarchy.GetCellData(intExpRow,intLinkCol)
																					Environment.Value("VAR_GetDoubleDependentCellValue")=intExpectedLoanNumber
																					UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Expected Cell Value is: "&intExpectedLoanNumber&" ", "Pass"
																			Else
																				UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Column is "&strFirstColName&" is not Found", "Fail"
																				Environment.Value("TestStepLog") = "False"
																				Environment.Value("TestObjectFlag") = "False"
																			End If		
																	Else
																			UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Column "&strSecColName&" Value  "&strExpectedText&" is not Found", "Fail"
																			Environment.Value("TestStepLog") = "False"
																			Environment.Value("TestObjectFlag") = "False"
																	End If
														Else
																		UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Column "&strSecColName&" is not Found", "Fail"
																		Environment.Value("TestStepLog") = "False"
																		Environment.Value("TestObjectFlag") = "False"
														End If
											Else
														UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Column  Value is "&strSearchText&" is not Found", "Fail"
														Environment.Value("TestStepLog") = "False"
														Environment.Value("TestObjectFlag") = "False"
											End If
									Else
												UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Column is "&strFirstColName&" is not Found", "Fail"
												Environment.Value("TestStepLog") = "False" 
												Environment.Value("TestObjectFlag") = "False"	
									End If
						Case "GetFirstRowCellValueOfCol"
									arrNumInputData = split(strData,";",-1,1) ' Getting the Count of Column to Verify
									For intNumCol = 0 to  Ubound(arrNumInputData)                            
										arrInputData=split(arrNumInputData(intNumCol),",",-1,1) ' Spliting the Column and ColumnValue
										strColName = arrInputData(0)
		'                                                                strExpectedColValue =arrInputData(1)
		'************************************Updated by Febin on 12/01/2011 for table with no: of columns changing dynamically********************
										intRowCount = ObjectHierarchy.RowCount
										For intRow = 1 to intRowCount - 1
												intColumnCount = ObjectHierarchy.ColumnCount(intRow)
												If intColumnCount > 1 Then
														intRCount= intRow
														Exit For
												End If
										Next
								
										'intColumnCount= ObjectHierarchy.ColumnCount(1)      
									'	**********Updated on 23 aug 11 by Suparna***************************
										'intRCount=1   
										'If intColumnCount=1 Then
											'	intColumnCount= ObjectHierarchy.ColumnCount(2)        
										'	intRCount=2
										'End If
									  '******************Update End		*****************
		'****************************************************************Febin's Update End******************************************************************
										   
										' To know  the Place of Column in the Table                                         
										For j=1 to  intColumnCount
												strCellData=ObjectHierarchy.GetCellData(intRCount,j)

												If instr(Trim(Ucase(strCellData )),Trim(Ucase(strColName)))Then         
																				intColNum = j   ' Getting the Column Number
																				isColFound = True
																				Exit For
												Else
																				isColFound=False            
												End If
										Next
										If isColFound = True Then
												'To Verify the Expected and Actual Column Value
												intRowCount = ObjectHierarchy.RowCount
												If intRowCount > 0 Then       
														strGetCellValue = ObjectHierarchy.GetCellData(intRCount+1,intColNum)
														Environment.Value("VAR_GetFirstRowCellValueOfCol") = Trim(strGetCellValue)
														UpdateReport "TESTSTEP", "","<font color=""green"">"& Environment.Value("strDescription") & "</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", " <font color=""green""> Column "& strColName &" is displayed with Value "& strGetCellValue  &" </font> ", "Pass"                                                                                             
												
												Else
														'No Rows Found For the particular search
														UpdateReport "TESTSTEP", "","<font color=""red"">"& Environment.Value("strDescription")& "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", " <font color=""red""> No Rows found for Column "& strColName &" and For Column Values "& strExpectedColValue &" </font> ", "Fail"         
														Environment.Value("TestStepLog") = "False"  
														Environment.Value("TestObjectFlag") = "False"
												End If
										Else        
												' Column is Not present in the table
												UpdateReport "TESTSTEP", "","<font color=""red"">"& Environment.Value("strDescription")& "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", " <font color=""red""> Column "& strColName &" Not Found </font> ", "Fail"         
												Environment.Value("TestStepLog") = "False"     
												Environment.Value("TestObjectFlag") = "False"
										End If                                                                    
								Next     					 	
					Case "SelectCheckBoxInTable"
								strFound="True"
								strFlag="False"
								strColStatus="True"
								strData = split(strData,",")
								strColumn1 = strData(0)
								strSearchText = strData(1)
								strColumn2 = strData(2)
								strStatus =  strData(3)
								'strSearchText = replace(strSearchText,",",", ")
								intColumnCount= ObjectHierarchy.ColumnCount(1)
								intRowCount=  ObjectHierarchy.rowcount       
							 ' To know  the Place of Column in the Table                                         
							   For j=1 to  intColumnCount
								   strCellData=ObjectHierarchy.GetCellData(1,j)
                                                                      'If instr(trim(strColumn1),Trim(strCellData))=1 Then  'Added by Suparan on 12 April:  if  table is having  one of the Column name  blank and if any other column name contains space it will excute true part.
									If trim(strColumn1)=Trim(strCellData) Then 
										For i=1 to intRowCount 
											strActualText=ObjectHierarchy.GetCellData(i,j)
											If Instr(strActualText, strSearchText) >= 1 Then
													intRowNumber=i
													strFound=True
													strFlag=True
													strColStatus=True
													 Exit for
											 Else											
													strFlag=False
											End if
										Next
										If strFlag= False Then      
											UpdateReport "TESTSTEP", "","<font color=""red"">"& Environment.Value("strDescription")& "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", " <font color=""red""> The Column : "& strColumn1 &" is not contain Expected Value <b><i>" & strSearchText & "</i></b></font> ", "Fail"         
											Environment.Value("TestStepLog") = "False"
											Environment.Value("TestObjectFlag") = "False"												
										End If
										Exit For
									Else
											strColStatus="False"
									End If  							
								Next
								If  strColStatus="False"Then
									UpdateReport "TESTSTEP", "","<font color=""red"">"& Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>","<font color=""red"">" & Time & "</font>"," <font color=""red""> The Column : <b><i>" & strColumn1 & "</i></b> is not Exist </font> ", "Fail"
									Environment.Value("TestObjectFlag") = "False"
									Environment.Value("TestStepLog") = "False"
															
								End If
								If  strFound= True Then
										 For k=1 to  intColumnCount
												strCellData=ObjectHierarchy.GetCellData(1,k)
												If Trim(Ucase(strColumn2)) = Trim(Ucase(strCellData)) Then 
													isColFound = True
													'ActualObject.ChildItem(intRowNumber,k,"Link",0).Click
													If  Instr(UCASE(strStatus),"CHECK")=1Then
															ObjectHierarchy.ChildItem(intRowNumber,k,"WebCheckBox",0).set "ON"
															UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Column <b>"& strColumn1 &"</b> with value <b>"& strSearchText  &"</b> check box is Checked Successfully.", "Done"  
															'UpdateReport "TESTSTEP", "","<font color=""green"">"& Environment.Value("strDescription") & "</font>", "<font color=""green"">" & StepStartTime & "</font>","<font color=""green"">" & Time & "</font>"," <font color=""green"">The Column <b>"& strColumn1 &"</b> with value <b>"& strSearchText  &"</b> check box is Checked Successfully. </font> ", "Pass"													
															Exit for										
													End If
													If  Instr(UCASE(strStatus),"UNCHECK")=1Then
															ObjectHierarchy.ChildItem(intRowNumber,k,"WebCheckBox",0).set "OFF"
															UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Column <b>"& strColumn1 &"</b> with value <b>"& strSearchText  &"</b> check box is unChecked Successfully.", "Done"  
															'UpdateReport "TESTSTEP", "","<font color=""green"">"& Environment.Value("strDescription") & "</font>", "<font color=""green"">" & StepStartTime & "</font>","<font color=""green"">" & Time & "</font>"," <font color=""green"">The Column <b>"& strColumn1 &"</b> with value <b>"& strSearchText  &"</b> check box is unChecked Successfully.</font> ", "Pass"								
															Exit For		
													End If												
												Else
													isColFound = False																		
												End If
										 Next
										 If isColFound =False Then						
												UpdateReport "TESTSTEP", "","<font color=""red"">"& Environment.Value("strDescription")& "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", " <font color=""red""> The Column : "& strColumn2 &"  not  Found </i></b></font> ", "Fail"         
												Environment.Value("TestStepLog") = "False"
												Environment.Value("TestObjectFlag") = "False"
										End If
								End If
					
			Case "StoreValueOfWebElementInTable"
											strFlag="True"
											strStatus="False"
											strExpectedField = strData
											'Descriptive Program
											Set Obj = Description.Create
											Obj("micclass").Value = "WebElement"
											Obj("html tag").Value = "TD"
											set objTablevalues= ObjectHierarchy.ChildObjects(obj)
											For intObjCount = 0 to objTablevalues.count-1
																			strActualField =  objTablevalues(intObjCount).GetRoProperty("innertext")
																			If Ucase(Trim(strActualField))=Ucase(trim(strExpectedField))Then
																											strExpectedFieldValue = objTablevalues(intObjCount +1).GetRoProperty("innertext")
																											set regEx = New RegExp
																											regEx.global = true
																											regEx.pattern = "[-?_,.()/:]"
																											strExpectedField = regEx.replace(strExpectedField, "")
																											strExpectedField = Trim(Replace(strExpectedField, " ", ""))
			'                                                                                               strExpectedField = Trim(Replace(strExpectedField, ".", ""))
			'                                                                                               strExpectedField = Trim(Replace(strExpectedField, ":", ""))
			'                                                                                               strExpectedField = Trim(Replace(strExpectedField, ")", ""))
			'                                                                                               strExpectedField = Trim(Replace(strExpectedField, "(", ""))
																											Environment.Value("VAR_GetValueFromTable"&strExpectedField)= Trim(strExpectedFieldValue)
																											strFound="True"          
																											UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Expected field <b>"& strExpectedField &"</b> value is <b>"& strExpectedFieldValue &"</b>", "Done"  
																							'               UpdateReport "TESTSTEP", "","<font color=""green"">"& Environment.Value("strDescription") & "</font>", "<font color=""green"">" & StepStartTime & "</font>","<font color=""green"">" & Time & "</font>"," <font color=""green"">The Expected field  <b>"& strExpectedField &"</b>  value is  <b>"& strExpectedFieldValue  &"</b> </font> ", "Pass"                                                                                                                                                                                                            
																											Exit For
																			End If                    
											Next
											If strFound="False" Then
																			UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Expected field  "& strExpectedField &" not found in table", "Fail"                  
															'UpdateReport "TESTSTEP", "","<font color=""red"">"& Environment.Value("strDescription")& "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", " <font color=""red""> The Expected field  "& strExpectedField &" Not Found in table </font>", "Fail"         
															Environment.Value("TestStepLog") = "False"   
															Environment.Value("TestObjectFlag") = "False"
											End If
											Set obj = Nothing
											Set objTablevalues = Nothing
											

					Case "ClickOnLinkBasedOnDependentCell"
							strSplitData = Split(strData,",")
							strDepCol = Trim(strSplitData(0))
							strDepColValue = Trim(strSplitData(1))
							strActCol = Trim(strSplitData(2))
							intRowCount = ObjectHierarchy.RowCount
							If InStr(strDepColValue,"fn_") > 0 OR InStr(strDepColValue, "VAR_") > 0 Then
									strDepColValue = Environment.Value(strDepColValue)
							End If
							For intRow = 1 to intRowCount - 1
									intColCount = ObjectHierarchy.ColumnCount(intRow)
									If intColCount > 1 Then
											intRCount = intRow
											Exit For
									End If
							Next
							If intRow = intRowCount Then
									UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime, Time, "The Table " & strObjectHier & " does not have valid columns", "Fail"
									Environment.Value("TestStepLog") = "False"
									Environment.Value("TestObjectFlag") = "False"
									Exit Function
							End If
							For intCol = 1 to intColCount
									strColHeader = Trim(ObjectHierarchy.GetCellData(intRCount,intCol))
									If InStr(UCASE(strColHeader),UCASE(strDepCol)) > 0 Then
											intDepCol = intCol
											isColFound = True
											Exit For
									Else
											isColFound = False
									End If
							Next
							If isColFound = True Then
									intActRow = ObjectHierarchy.GetRowWithCellText(strDepColValue)
									If intActRow <> -1 Then
											isRowFound = True
									Else
											isRowFound = False
									End If
							Else
									UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime,  Time ,"The dependant column " & strDepCol & " is not found in the Table " & strObjectHier, "Fail"
									Environment.Value("TestStepLog") = "False"
									Environment.Value("TestObjectFlag") = "False"
									Exit Function
							End If
							If isRowFound = True Then
									For intCol = 1 to intColCount
											strActColHeader = Trim(ObjectHierarchy.GetCellData(intRCount,intCol))
											If  UCASE(strActColHeader) = UCASE(strActCol) Then
													isActColFound = True
													intActCol = intCol
													Exit For
											Else
													isActColFound = False
											End If
									Next
							Else
									UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime,  Time ,"The dependant row value " & strDepColValue & " is not found in the Table " & strObjectHier, "Fail"
									Environment.Value("TestStepLog") = "False"
									Environment.Value("TestObjectFlag") = "False"
									Exit Function
							End If
							If isActColFound = True Then
								'Updated by suparna on 21 March: If 2 links are present and required to click on 2nd link , need to pass one more parameter value i,e index of 2nd link
								If ubound(strSplitData)=3 Then
										If ObjectHierarchy.ChildItem(intActRow,intActCol,"Link",strSplitData(3)).Exist(5) Then			' Added 5 seconds to overcome the sync issues, Shrinidhi 22nd Mar
												ObjectHierarchy.ChildItem(intActRow,intActCol,"Link",strSplitData(3)).Click
												UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime,  Time ,"The link in the column " & strActCol & " with the dependant column " & strDepCol & " and value "& strDepColValue & " is clicked", "Done"
										Else
												UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime,  Time ,"The link in the column " & strActCol & " with the dependant column " & strDepCol & " and value "& strDepColValue & " is not available", "Fail"
												Environment.Value("TestStepLog") = "False"
												Environment.Value("TestObjectFlag") = "False"
										End If
								Else
										If ObjectHierarchy.ChildItem(intActRow,intActCol,"Link",0).Exist(5) Then							'Added 5 seconds to overcome the sync issues, Shrinidhi 22nd Mar
											ObjectHierarchy.ChildItem(intActRow,intActCol,"Link",0).Click
											UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime,  Time ,"The link in the column " & strActCol & " with the dependant column " & strDepCol & " and value "& strDepColValue & " is clicked", "Done"
									Else
											UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime,  Time ,"The link in the column " & strActCol & " with the dependant column " & strDepCol & " and value "& strDepColValue & " is not available", "Fail"
											Environment.Value("TestStepLog") = "False"
											Environment.Value("TestObjectFlag") = "False"
									End If	
								End If

								
							Else
									UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime,  Time ,"The actual column " & strActCol & " is not found in the Table " & strObjectHier, "Fail"
									Environment.Value("TestStepLog") = "False"
									Environment.Value("TestObjectFlag") = "False"
									Exit Function
							End If	
					End Select
	Else
		UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Table does not exist</font>", "Fail"
		Environment.Value("TestObjectFlag") = "False"
		Environment.Value("TestStepLog") = "False"
	End If
End Function

'=====================================================================================================================================================================================
' Name of the Function     			  : fn_VerifyClearOption
' Description       		   		 	     : This function is use to Verify that  Particular Object should not contain any Value.
' Date and / or Version       	    : 13-05-2011
' Author									      : Manish
' Input Parameters					: ScreenName and Object Only
' Example Call							 : Call fn_VerifyClearOption()
'Return    Value                        : no Return Value      

'======================================================================================================================================================================================

Function fn_VerifyClearOption(strScenarioName,strData)
		StepStartTime = Time
		strObjectHierarchy = Datatable.value("APP_SCREEN_NAME", strScenarioName)              
		strObject = Datatable.Value("OBJECT", strScenarioName)
		Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
		If strObject <> ""  Then
				strActualObject = strObjectHierarchy & "." & strObject
		ElseIf strObject = "" Then
				strActualObject = strObjectHierarchy
		End If
		Set ActualObject = Eval(strActualObject)
		Call fnc_wait(ActualObject)  
		If  Instr(strObject,"WebList")>=1  Then
				strfieldValue = ActualObject.GetROProperty("selected item index")
				If strfieldValue = 0 Then
						UpdateReport "TESTSTEP", "","<font color=""green"">"& Environment.Value("strDescription") & "</font>", "<font color=""green"">" & StepStartTime & "</font>","<font color=""green"">" & Time & "</font>"," <font color=""green""> Dropdown field value is cleared as expected</font> ", "Pass" 
				Else				
						UpdateReport "TESTSTEP", "", "<font color=""red"">"& Environment.Value("strDescription")&"</font>",  "<font color=""red"">"& StepStartTime &"</font>", "<font color=""red"">"& Time&"</font>" , "<font color=""red"">Dropdown field value is not cleared as Expected</font> ", "Fail"         
						Environment.Value("TestStepLog") = "False"
				End If
		Elseif Instr(strObject,"WebEdit")>=1 then
				strfieldValue = ActualObject.GetROProperty("Value")
				If strfieldValue = "" Then
						UpdateReport "TESTSTEP", "","<font color=""green"">"& Environment.Value("strDescription") & "</font>", "<font color=""green"">" & StepStartTime & "</font>","<font color=""green"">" & Time & "</font>"," <font color=""green"">Web Edit field value is cleared as expected</font> ", "Pass" 
				Else		
						UpdateReport "TESTSTEP", "", "<font color=""red"">"& Environment.Value("strDescription")&"</font>",  "<font color=""red"">"& StepStartTime &"</font>", "<font color=""red"">"& Time&"</font>" , "<font color=""red"">Web Edit field value is not cleared as expected</font> ", "Fail"            		
						Environment.Value("TestStepLog") = "False"
				End If  
		End If
		
End function
'=====================================================================================================================================================================================
' Name of the Function     			  : fn_StoreValueinTempVar
' Description       		   		 	     : This function is used to store a dynamic value to the user defined variable so as to reuse the variable whenever required
' Date and / or Version       	    : 12-01-2011
' Author									      : Febin Mathew
' Input Parameters					: 	Variable Name, Variable Value
' Example Call							 : Call fn_StoreValueinTempVar(VAR_Temp,123))
'Return    Value                        : no Return Value      

'======================================================================================================================================================================================
Function fn_StoreValueinTempVar(strDataVal)
		StepStartTime = Time
		strData = strDataVal		
		strData = strDataVal
		strSplitData = Split(strData,",")
		strVarName = Trim(strSplitData(0))
		strVarValue = Trim(strSplitData(1))
		If Instr(strVarValue, "fn_") > 0 OR Instr(strVarValue,"VAR_") > 0 Then
				'If Environment.Value(strVarValue) <> NULL Then
						strVarValue = Environment.Value(strVarValue)
				'Else
'						UpdateReport "TESTSTEP", "", "Store temporary value",StepStartTime, Time,"The Environment Variable " & strVarValue & " is not available as expected", "Fail"            		
'						Environment.Value("TestStepLog") = "False"
'						Environment.Value("TestObjectFlag") = "False"
'						Exit Function
'				End If
		End If
		Environment.Value(strVarName) = strVarValue
End Function

'=====================================================================================================================================================================================
' Name of the Function     			  : fn_CompareVariableValues
' Description       		   		 	      : This function is used to compare two variable and give the result
' Date and / or Version       	    : 12-01-2011
' Author									      : Febin Mathew
' Input Parameters					: 	Actual value, Expected value
' Example Call							 : Call fn_StoreValueinTempVar(123,fn_Variable))
'Return    Value                        : no Return Value
'======================================================================================================================================================================================

Function fn_CompareVariableValues(strScenarioName,strDataVal)
		StepStartTime = Time
		strSplitData=Split(strDataVal,",")
		strActValue = strSplitData(0)
		strExpValue = strSplitData(1)
		If Instr(strActValue, "fn_") > 0 OR Instr(strActValue,"VAR_") > 0 Then
				'If Environment.Value(strActValue) <> NULL Then
						strActValue = Environment.Value(strActValue)
				Else
				strActValue = strActValue
'						UpdateReport "TESTSTEP", "", "Compare 2 variable values",  "", "" , "The Environment Variable " & strActValue & " is not available as expected", "Fail"            		
'						Environment.Value("TestStepLog") = "False"
'						Exit Function
'				End If
		End If
		If Instr(strExpValue, "fn_") > 0 OR Instr(strExpValue,"VAR_") > 0 Then
				'If Environment.Value(strExpValue) <> NULL Then
						strExpValue = Environment.Value(strExpValue)
				Else
				strExpValue =strExpValue
'						UpdateReport "TESTSTEP", "", "Compare 2 variable values",  "", "" , "The Environment Variable " & strExpValue & " is not available as expected", "Fail"            		
'						Environment.Value("TestStepLog") = "False"
'						Exit Function
'				End If
		End If
		If Instr(Trim(Ucase(strActValue)),Trim(Ucase(strExpValue)))>=1 Then
				UpdateReport "TESTSTEP", "","Compare 2 variable values", StepStartTime, Time, "Expected  value- <b> " & strExpValue &" </b> and Actual value- <b>" & strActValue &" </b> are matching", "Pass"
		Else
				UpdateReport "TESTSTEP", "","Compare 2 variable values", StepStartTime, Time, "Expected  value- <b> " & strExpValue &" </b> and Actual value- <b>" & strActValue &" </b> does not  match", "Fail"
				Environment.Value("TestStepLog") = "False"
		  End If
End Function

'==================================================================================================================================================
' Name of the Function    	  : fn_SelectItemContainingValue
' Description                          : 
' Date and / or Version 	  :  12/11/2011
' Author                               		:  Vignesh
' Input Parameters              : 
' Example Call                     : 
'================================================================================================================================================f=
Function fn_SelectItemContainingValue(strScenarioName, strDataVal)
   'For log
	StepStartTime = Time
	'Copy of parameters
	strData = strDataVal		
	'Get the application screen name i.e. Browser and page
	strObjectHierarchy = Datatable.value("APP_SCREEN_NAME", strScenarioName)		' Object Hierarchy value ex: Browser("Login").Page("Login")
	'Get the object to be worked on
	strObject = Datatable.Value("OBJECT", strScenarioName)
	'Get the description, used for reporting
	Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)

	'Combining the Browser, page and Object
	If strObject <> ""  Then
			strActualObject = strObjectHierarchy & "." & strObject
	ElseIf strObject = "" Then
			strActualObject = strObjectHierarchy
	End If

	'Making the string object hierarchy into an object
	Set ActualObject = Eval(strActualObject)
	'Sync the application till the required object is found
	Call fnc_wait(ActualObject)

   arrAllItems = Split(ActualObject.GetROProperty("all items"), ";")
	For intItemPos = 0 to Ubound(arrAllItems)
		If Instr(arrAllItems(intItemPos), strData) > 0 Then
			ActualObject.Select "#" &intItemPos+1
			UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime, Time,"Item containing value " &strData &" selected from drop down" , "Done"
			Exit Function
		End If
  Next
  UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime, Time,"Item not found in the drop down." , "Fail"
  Environment.Value("TestStepLog") = "False"
  Environment.Value("TestObjectFlag") = "False"
End Function


'==================================================================================================================================================
' Name of the Function               : fn_ClickDialogBoxButton
' Description                  : This function is used  to click on a button if dialog box appeares
' Date and / or Version                       : 06-07-2011
' Author                                                                                                                                                      : Suparna
' Input Parameters                                                                          : NA
' Example Call                                                                                                     : Call fn_ClickDialogBoxButton()                                                            : 
'==================================================================================================================================================
Function fn_ClickDialogBoxButton(strScenarioName, strDataVal)
	StepStartTime = Time
	strData = strDataVal
	strObjectHierarchy = Datatable.value("APP_SCREEN_NAME", strScenarioName)                              ' Object Hierarchy value ex: Browser("Login").Page("Login")
	strObject = Datatable.Value("OBJECT", strScenarioName)
	Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
	
	If strObject <> ""  Then
		strActualObject = strObjectHierarchy & "." & strObject
	ElseIf strObject = "" Then
		strActualObject = strObjectHierarchy
	End If
	
	Set ActualObject = Eval(strActualObject)

	If  ActualObject.Exist(1) Then
		ActualObject.Click
	End If
End Function

'==================================================================================================================================================
' Name of the Function                                                    : fn_ClickonMsgboxButton
' Description                                                                                            : This function is used  to click on msgbox button
' Date and / or Version                       : 
' Example Call                                                                                                     : fn_ClickonMsgboxButton(brwname)
'==================================================================================================================================================
Function fn_ClickonMsgboxButton(strScenarioName, strDataVal)
					StepStartTime = Time
					strBrowserName = strDataVal 'Datatable.Value("INPUTDATA_PARAMETER", strScenarioName)
					If Browser(strBrowserName).Exist(3) Then
								If Browser(strBrowserName).Dialog("Class Name:= Dialog", "nativeclass:= #32770").exist(1) then                                                                                      
																Set obj= description.Create
																obj("micclass").value="WinButton"
																set Objbtn=Browser(strBrowserName).Dialog("Class Name:= Dialog", "nativeclass:= #32770").ChildObjects(obj)
																If Objbtn.count>0 then
																					Objbtn(0).click
																					Exit Function
																					'UpdateReport "TESTSTEP", "",Environment.Value("strDescription") , StepStartTime, Time, "The Dialog box button is clicked", "Pass"
																End if
																Call 	fn_ClickonMsgboxButton	(strScenarioName, strDataVal)
								End if
							'Browser(strBrowserName).Close																																																																																																																																																																																																																																																																						
					End if		
									
End function


'==================================================================================================================================================
' Name of the Function     			  : OperateOnWinEdit
' Description       		   		 	     : This function is used to enter or retrieve the values on / from the WebEdit object
' Date and / or Version       	    : 
' Example Call							 : OperateOnWinEdit("Browser(""TIPS"").Page(""TitleGrid"")","LoginName", "SETVALUE", "Test123")
'==================================================================================================================================================
'=========================================================================================================================================
Function OperateOnWinEditor(strObjectHierarchy, strObjectName, strAction_Name, strData)
   
		If  strAction_Name <> "STOREVALUE" AND Instr(strData, "VAR_") = 1 Then				' If the value to be taken from already saved variable
			strData = Environment.Value(strData)
		End If

		If  Instr(strData, "fn_") = 1 Then
			strData = Environment.Value(strData)
		End If

		Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.WinEditor(strObjectName)
		StepStartTime = Time
        If fnc_wait(ActualObject) = "True" Then
				Select Case strAction_Name
					  Case "SETVALUE"	
	  						  Wait(1)
	  						  ActualObject.Click
							  ActualObject.Set strData							
							  If Instr(Environment.Value("strDescription"), "Password") > 0 Then
									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Password is entered", "Done"
							  Else
									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Value: <font color=""blue""><b><i>" & strData & "</i></b></font> is entered in field <b>" & Mid(strObjectName,4,Len(strObjectName)-3) &"</b> successfully", "Done"
							  End If
							  
					   Case "CLICK"
							  ActualObject.Click									
							  UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The WebEdit: <b>" & strObjectName & "</b> is clicked", "Done"
							
					  Case "CLEARFIELD"
					  		  ObjectHierarchy.Activate
							  Set wclr = CreateObject("Wscript.Shell")  
							  wclr.SendKeys("{DELETE}") 									
							  UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The WebEdit: <b>" & strObjectName & "</b> value is cleared", "Done"
							
					  
					  Case "TYPEVALUE"
							  Wait(1)
  							  ActualObject.Click
							  ActualObject.Type strData
							   If Instr(Environment.Value("strDescription"), "Password") > 0 Then
									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Password is entered", "Done"
							  Else
									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Value: <font color=""blue""><b><i>" & strData & "</i></b></font> is entered in field <b>" & Mid(strObjectName,4,Len(strObjectName)-3) &"</b> successfully", "Done"
							  End If
					   Case "COMPAREVALUE"
							   If Instr(strData,"fn_") >0 OR Instr(strData,"VAR_") >0  Then
									strExp = Environment.Value(strData)
							  Else
									strExp = strData
							  End If	
							  strActual = ObjectHierarchy.WinEditor(strObjectName).GetROProperty("text")
							  For i = 1 To len(strActual) Step 1
								 strTemp = Asc(Mid(strActual,i,1))
									If (strTemp >= 65 AND strTemp <= 95) OR (strTemp >= 97 AND strTemp <= 122) Then
										strActualData =strActualData & Mid(strActual,i,1)
									End If
								Next
							  If Trim(UCASE(strExp)) = Trim(UCASE(strActualData)) Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Expected Value - "& strExp &" and Actual value - "& strActualData & " are matching</font>", "Pass"
							  Else
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Value mismatch, Expected Value is: <i>" & strExp & "</i>, and Actual value on the application is: <i>" & strActualData & "</i></font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If		  
					  
					  Case "CHECKEXIST"
							  strExp = strData
							  strActual = ObjectHierarchy.WinEditor(strObjectName).Exist
							  If UCASE(strExp) = UCASE(CStr(strActual)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object exists as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strActual)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object does not exist</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  ElseIf UCASE(strExp) = UCASE(CStr(strActual)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object does not exist as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strActual)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object exists</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
					 Case "SETPASSWORD"
							 ActualObject.SetSecure strData
							 UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Password is entered", "Done"
					 Case "STOREVALUE"
							 strVariableName = strData
							 Environment.Value("VAR_"& strVariableName) = ObjectHierarchy.WinEditor(strObjectName).GetROProperty("text")
							 'msgbox  Environment.Value("VAR_"& strVariableName)
							 UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The <i>"& strData &"</i> value: <i>"& Environment.Value("VAR_"& strVariableName)& "</i> is stored", "Done"
					 Case "CHECKENABLED"
								blnObjDisable= ObjectHierarchy.WinEditor(strObjectName).GetROProperty("enabled")								
								If blnObjDisable = "1" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Disabled as expected</font>", "Pass"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Enabled as expected</font>", "Pass"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Enabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								ElseIf blnObjDisable = "1" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Disabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								End If
						End Select
        Else
                UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The WinEditor - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If

End Function



'==================================================================================================================================================
' Name of the Function     			  : OperateOnWinButton
' Description       		   		 	     : This function is used to enter or retrieve the values on / from the Web object
' Date and / or Version       	    : 
' Example Call							 : OperateOnWinButton("Browser(""TIPS"").Page(""TitleGrid"")","OK", "CLICK", "")
'==================================================================================================================================================
Function OperateOnWinButton(strObjectHierarchy, strObjectName, strAction_Name, strData)
   
		Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.WinButton(strObjectName)
		StepStartTime = Time
		Call fnc_wait(ActualObject)
		
		If fnc_wait(ActualObject) = "True" Then
					Select Case strAction_Name
					  Case "CLICK"
							  On Error Resume Next
							  Wait(1)
							  ObjectHierarchy.WinButton(strObjectName).Click
							  If Err.Number = 0 Then
							  	 UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Button: <b>" & Mid(strObjectName,4,Len(strObjectName)-3) & "</b> is clicked successfully", "Done"
							  Else
							  	 UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The button not clicked Due to error " & Err.Description & "</font>", "Fail"
 							 	 Environment.Value("TestStepLog") = "False"
 							 End If
							  
					  Case "CHECKEXIST"
							  strExp = strData
							  strAct = ActualObject.Exist(0)
							  If UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object exists as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object does not exist</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  ElseIf UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object does not exist as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object exists</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
						Case "CHECKENABLED"
							  blnObjDisable= ActualObject.GetROProperty("disabled")
							  '**Start***Updated by Manish on 6/29/11***									
								If blnObjDisable = "1" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Disabled as expected</font>", "Pass"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Enabled as expected</font>", "Pass"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Enabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								ElseIf blnObjDisable = "1" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Disabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								End If
						End Select
	           Else
							UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Button - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
							'UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Button - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
							Environment.Value("TestObjectFlag") = "False"
							Environment.Value("TestObjectFlag") = "False"
			End If		
End Function

'==================================================================================================================================================
' Name of the Function     			  : OperateOnWinComboBox
' Description       		   		 	     : This function is used to select or retrieve the values from the WebList object
' Date and / or Version       	    : 
' Example Call							 : OperateOnWinComboBox("Browser(""TIPS"").Page(""TitleGrid"")","StatusCode", "SELECTVALUE", "Test")
'==================================================================================================================================================
Function OperateOnWinComboBox(strObjectHierarchy, strObjectName, strAction_Name, strData)

On Error Resume Next

		If  strAction_Name <> "STOREVALUE" AND Instr(strData, "VAR_") = 1 Then				' If the value to be taken from already saved variable
				strData = Trim(Environment.Value(strData))
		End If

	    If  Instr(strData, "fn_") = 1Then
			strData = Environment.Value(strData)
		End If
			
	    Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.WinComboBox(strObjectName)
		StepStartTime = Time
        If fnc_wait(ActualObject) = "True" Then
				Select Case strAction_Name
					  Case "SELECTVALUE"
					  		 Wait(1)
'							  ActualObject.Click
							  ObjectHierarchy.WinComboBox(strObjectName).Select strData							 
							  strActualValue = ObjectHierarchy.WinComboBox(strObjectName).GetRoProperty("selection")
							  strActualValue = Trim(strActualValue)
							  If strComp(strActualValue,strData) <> 0 Then
									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value <font color=""blue""><b><i>" & trim(strData) & "</i></b></font> is not available in the Dropdown <b>"& Mid(strObjectName,4,Len(strObjectName)-3)&"</b>", "Fail"
									Environment.Value("TestStepLog") = "False"
									Environment.Value("TestObjectFlag") = "False"
							  Else
									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value <font color=""blue""><b><i>" &  trim(strData) & "</i></b></font> is selected from the Dropdown <b>"& Mid(strObjectName,4,Len(strObjectName)-3)&"</b>", "Done"
							  End If
					  Case "COMPAREVALUE"
								  If instr(strData,"fn_") >0 OR Instr(strData,"VAR_") >0  Then
										strExp = Environment.Value(strData)
								  Else
										strExp = strData
								  End If							  
								  strActual = ObjectHierarchy.WinComboBox(strObjectName).GetROProperty("text")
								  If Trim(UCASE(strExp)) = Trim(UCASE(strActual)) Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Expected Value - "& strExp &" and Actual value - "& strActual & " are matching</font>", "Pass"
								  Else
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Value mismatch, Actual value is: <i>" & strActual & "</i></font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								  End If
					  Case "STOREVALUE"
								 strVariableName = strData
								 Environment.Value("VAR_"&strVariableName) = ObjectHierarchy.WinComboBox(strObjectName).GetROProperty("value")	
								 UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The "& strData &" value is stored</font>", "Done"
					  Case "CHECKEXIST"
								  strExp = strData
								  strAct = ObjectHierarchy.WinComboBox(strObjectName).Exist
								  If UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object exists as expected</font>", "Pass"
								  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object does not exist</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								  ElseIf UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object does not exist as expected</font>", "Pass"
								  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object exists</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								  End If
						Case "CHECKENABLED"
									blnObjDisable= ObjectHierarchy.WinComboBox(strObjectName).GetROProperty("enabled")					
									If blnObjDisable = "1" AND UCASE(strData) = "FALSE" Then
											UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Disabled as expected</font>", "Pass"
									ElseIf blnObjDisable = "0" AND UCASE(strData) = "TRUE" Then
											UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Enabled as expected</font>", "Pass"
									ElseIf blnObjDisable = "0" AND UCASE(strData) = "FALSE" Then
											UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Enabled</font>", "Fail"
											Environment.Value("TestStepLog") = "False"
									ElseIf blnObjDisable = "1" AND UCASE(strData) = "TRUE" Then
											UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Disabled</font>", "Fail"
											Environment.Value("TestStepLog") = "False"
									End If
				End Select
	   Else
                UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The WebList - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If
End Function

'==================================================================================================================================================
' Name of the Function     			  : OperateOnWinCheckBox
' Description       		   		 	     : This function is used to select or retrieve the values from the WebCheckBox object
' Date and / or Version       	    : 
' Example Call							 : OperateOnWinCheckBox("Browser(""TIPS"").Page(""TitleGrid"")","SELECT", "CHECK", "")
'==================================================================================================================================================
Function OperateOnWinCheckBox(strObjectHierarchy, strObjectName, strAction_Name, strData) 

	    Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.WinCheckBox(strObjectName)
		StepStartTime = Time
		If fnc_wait(ActualObject) = "True" Then
				Select Case strAction_Name
						  Case "CHECK"
								If  ActualObject.GetROProperty("checked") = "OFF" Then
'										ActualObject.Set "ON"
										Wait(1)
										ActualObject.Click
										UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Check box <b>" & Mid(strObjectName,4,Len(strObjectName)-3)&"</b> is checked successfully", "Done"
								Else
										UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Check box <b>" & Mid(strObjectName,4,Len(strObjectName)-3)&"</b> is checked successfully", "Done"
								End If
						  Case "UNCHECK"
								If  ActualObject.GetROProperty("checked") = "ON" Then
'										ActualObject.Set "OFF"
										ActualObject.Click
										UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Check box <b>" & Mid(strObjectName,4,Len(strObjectName)-3)&"</b> is unchecked successfully", "Done"
								 Else
										UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Check box <b>" & Mid(strObjectName,4,Len(strObjectName)-3)&"</b> is unchecked successfully", "Done"
								End If
'***************************** Added by Suresh on 28/07/2011 *************************************************
						Case "CHECKEXIST"
							  strExp = strData
							  strAct = ActualObject.Exist
							  If UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object exists as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object does not exist</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  ElseIf UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object does not exist as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object exists</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
						Case "CHECKENABLED"
								blnObjDisable= ActualObject.GetROProperty("disabled")
                            	If blnObjDisable = "1" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Disabled as expected</font>", "Pass"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Enabled as expected</font>", "Pass"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Enabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								ElseIf blnObjDisable = "1" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Disabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								End If
								Case "CHECKCHECKED"
							 			blnObjDisable= ActualObject.GetROProperty("checked")
                            	If blnObjDisable = "1" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Checked as expected</font>", "Pass"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Not Checked as expected</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Not Checked as expected</font>", "Pass"
                               	ElseIf blnObjDisable = "1" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Checked</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								End If
'**************************************************************************************************************************************
				End Select
		Else
				UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Checkbox - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If
End Function


'==================================================================================================================================================
' Name of the Function     			  : OperateOnWinRadioButton
' Description       		   		 	     : This function is used to select or retrieve the values from the WinRadioButton object
' Date and / or Version       	    : 
' Example Call							 : OperateOnWinRadioButton("Browser(""TIPS"").Page(""TitleGrid"")","SELECT", "CHECK", "")
'==================================================================================================================================================
Function OperateOnWinRadioButton(strObjectHierarchy, strObjectName, strAction_Name, strData) 

	    Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.WinRadioButton(strObjectName)
		StepStartTime = Time
		If fnc_wait(ActualObject) = "True" Then
				Select Case strAction_Name
						  Case "SELECT"
		  								Wait(1)
										ActualObject.Set
										UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "RadioButton <b>"& Mid(strObjectName,4,Len(strObjectName)-3) & "</b> is selected successfully", "Done"
				End Select
		Else
				UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The RadioButton - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If
End Function

'==================================================================================================================================================
' Name of the Function     			  : OperateOnWinMenu
' Description       		   		 	     : This function is used to select or retrieve the values from the WinRadioButton object
' Date and / or Version       	    : 
' Example Call							 : OperateOnWinMenu("", "Save", "")
'==================================================================================================================================================
Function OperateOnWinMenu(strObjectHierarchy, strObjectName, strAction_Name, strData) 
		On Error Resume next
		StepStartTime = Time
	    Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.WinMenu(strObjectName)
		ActualObject.Select strData
		
		If Instr(strData, "<")=1 OR  Instr(strData,">")=1 Then
			
			strData = Replace(Replace(strData, "<",""),">","")
			If Instr(1,strData,";") <> 0 Then
				arrData = Split(strData, ";")
				Menu = arrData(0)
				SubMenu = arrData(1)
				UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"The SubMenu Value" & "<font color=""blue""><b><i>'"& SubMenu & "'</i></b></font> is selected successfully from Menu Tab" & "<font color=""blue""><b><i>'"& Menu & "'</i></b></font> in the Window", "Done"
			ElseIf Instr(1,strData,";") = 0 Then
				Menu = strData
				UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"The Menu Value" & "<font color=""blue""><b><i>'"& Menu & "'</i></b></font> is selected successfully from Menu Tab in the Window", "Done"
			Else
				UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>","The Menu Value" & "<font color=""red""><b><i>'"& Menu & "'</i></b></font> is not selected successfully  in the Window", "Fail"
				Environment.Value("TestStepLog") = "False"
			End If	
		End If
		
End Function



'*******************************************************************************************************************************************************************************
' Name of the Function     			 : fn_ObjectIdentification
' Description       		   		 	      : This function validate web Element
' Date and / or Version       	    : 6/10/2013
' Author									      : Srivaths
' Input Parameters					  : None 
' Example Call							 : Call fn_ObjectIdentification(wdwWindow,strMicClass,strHTMLTag,TextToCompare,ActionToBePerformed)

'*******************************************************************************************************************************************************************************
Function fn_ObjectIdentification(strScenarioName,ByVal strDataVal)
		
		On Error Resume Next
		
		arrDataVal = Split(strDataVal,",")
		strMicClass = arrDataVal(0)
		strHTMLTag = arrDataVal(1)
		TextToCompare = arrDataVal(2)
		ActionToBePerformed = arrDataVal(3)

		'Replacing # character with comma
		If instr(1,TextToCompare,"#") <> 0 Then
	    	TextToCompare=replace(TextToCompare,"#",",")
	    End If
	    
	    'Replacing $ character with comma
		If instr(1,TextToCompare,"~") <> 0 Then
	    	TextToCompare=replace(TextToCompare,"~","""")
	    End If
	    
	    
		StepStartTime = Time
		wdwWindow = Datatable.Value("APP_SCREEN_NAME", strScenarioName)              
		Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
		
		intFailCount = 0
		intPassCounter = 0
		Set oReqObject = Description.Create()
		oReqObject("micclass").Value = strMicClass
		oReqObject("html tag").Value = strHTMLTag
		Set wdwWindow = Eval(wdwWindow)
		
		Wait(1)
		
		intHeaderChildObjectCount = wdwWindow.ChildObjects(oReqObject).Count
		Set objHeaderColl = wdwWindow.ChildObjects(oReqObject)
		
		If Instr(TextToCompare,"||")>0 Then
			arrTextToCompare = Split(TextToCompare,"||")
		Else
			Dim arrTextToCompare(0)
			arrTextToCompare(0) = TextToCompare
		End If
		
		For intIterator = 0 to intHeaderChildObjectCount				'For loop Block 1
		
			For intCounter = 0 to Ubound(arrTextToCompare)					'For Loop Block 2
		
				If Instr(Trim(objHeaderColl(intIterator).GetRoProperty("innertext")),arrTextToCompare(intCounter))>0 Then ' Comparision Block for Object Identification
					'To display in the Result Log to correct the exact string
					strExpMessageToCompare=arrTextToCompare(intCounter)
					intPassCounter = intPassCounter + 1
					intClickCounter = intIterator
					If intPassCounter = 1 Then
						strDataToBeDisplayed = arrTextToCompare(intCounter)
					elseIf intPassCounter > 1 Then
						strDataToBeDisplayed = strDataToBeDisplayed + ";" + arrTextToCompare(intCounter)
					End If
					If (intPassCounter = Ubound(arrTextToCompare) + 1) then
						strMessageToCompare = Trim(objHeaderColl(intIterator).GetRoProperty("innertext"))
						Exit For
					End If
				else
					Exit For
				End If 'End Block for Object Identification
					
			Next 'End Of For loop Block 1
			
			 If (intPassCounter = Ubound(arrTextToCompare) + 1) then
			 	Exit For
			 End If
			 
		Next 'End Of For loop Block 2
		
		If UCASE(ActionToBePerformed) = "VALIDATE" and (intPassCounter = Ubound(arrTextToCompare) + 1) then
			UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime, Time, "<font color=""blue""><b><i>'" & strExpMessageToCompare & "'</i></b></font> displayed successfully", "Done"
		ElseIf UCASE(ActionToBePerformed) = "VERIFYINMESSAGE" and (intPassCounter = Ubound(arrTextToCompare) + 1) then
			UpdateReport "TESTSTEP", "","<b>'" & strDataToBeDisplayed & "'</b> should be present in " & strExpMessageToCompare, StepStartTime, Time, "<font color=""blue""><b><i>'" & strDataToBeDisplayed & "'</i></b></font> displayed successfully", "Done"
		ElseIf UCASE(ActionToBePerformed) = "ERRORMESSAGE" and (intPassCounter = Ubound(arrTextToCompare) + 1) then
			If strMessageToCompare <> "" Then
				UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime, Time,"The Error Message - " & "<font color=""blue""><b><i>'" & strExpMessageToCompare & "'</i></b></font> was displayed successfully", "Pass"
			else
				UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime, Time,"Error message - " & "<font color=""blue""><b><i>'" & strExpMessageToCompare & "'</i></b></font> is not displayed", "Fail"
				intFailCount = intFailCount + 1
			End If
		ElseIf UCASE(ActionToBePerformed) = "CLICK" and (intPassCounter = Ubound(arrTextToCompare) + 1) Then 'Block to perform Click Operation
			objHeaderColl(intClickCounter).Click
			UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime, Time, "Entity clicked successfully", "Done"
		Else
			UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime, Time, "Object Not Found", "Fail"
			intFailCount = intFailCount + 1
		End If
		
		If intFailCount = 0 and (intPassCounter = Ubound(arrTextToCompare) + 1) Then
			fn_ObjectIdentification = True
		Else
			fn_ObjectIdentification = False
		End If
		
End Function
'==================================================================================================================================================
' Name of the Function                    : fn_ValidateEmailInOutlook
' Description                                         : This Function to open outlook, access inbox folder and read e-mail and click on the required link from inbox
' Date and / or Version                     : 6/10/2013
' Example Call                                       : fn_ValidateEmailInOutlook(strSubject,strScenario,strActionToBePerformed,strValue)
'==================================================================================================================================================

Function fn_ValidateEmailInOutlook(strScenarioName,strDataVal)

arrDataVal = Split(strDataVal,",")
strSubject = arrDataVal(0)
strScenario = arrDataVal(1)
strActionToBePerformed = arrDataVal(2)
strValue = arrDataVal(3)

SystemUtil.Run "\\corp\corpdata\MyAccount_E2E_Testing_Automation\E2E Automation\Automate\Initialization\OutlookSecurity.vbs"

srtDescription = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
StepStartTime = Time
strBodyContent = fn_ReadEmailFromOutlook(strSenderdetails,strSubject,srtDescription)
blnClickLink = False

strSenderdetails = DataTable.Value("GParam_SenderDetails","GLOBALPARAMETERS")
strURL = DataTable.Value("GParam_RequiredURL","GLOBALPARAMETERS")

Select Case Ucase(strActionToBePerformed)

                Case "CLICK"
                                arrReqURL = Split(strBodyContent,"HYPERLINK """)
                                For intCounter = 0 to Ubound(arrReqURL)
                                                If Instr(arrReqURL(intCounter),strURL) > 0 Then
                                                                intReqURLend = Instr(arrReqURL(intCounter),strScenario)
                                                                strReqURl = Mid(arrReqURL(intCounter),1,intReqURLend-2)
                                                                If inStr(srtDescription,"MyAccountMobile") > 0 then
                                                                	intSocalgas = Instr(strReqURL,".socalgas.com")
																	strReqpartURL = Mid(strReqURL,1,intSocalgas)
																	intTokenId = Instr(strReqURL,"TokenId=")
																	intTokenId = intTokenId + 8
																	strTokenId = Mid(strReqURL,intTokenId ,intReqURLend-2)
																	If Instr(strReqURL,"forgotPassword")>0 Then
	                                                                	strReqURl = strReqpartURL&"socalgas.com/forgotpassword/ValidateTokenServlet?" & "TokenId=" & strTokenId &"&ms=Y"
                                                                	ElseIf Instr(strReqURL,"forgotBoth")>0 Then
                                                                		strReqURl = strReqpartURL&"socalgas.com/forgotboth/ValidateTokenServlet?" & "TokenId=" & strTokenId &"&ms=Y"
                                                                	End if
                                                                End If
                                                                UpdateReport "TESTSTEP", "",srtDescription, StepStartTime, Time, "<b>'" & strScenario & "'</b> link is clicked successfully", "Done"
                                                                fn_ValidateEmailInOutlook = True
                                                                blnClickLink = True
                                                                fn_LauchURL strScenarioName, strReqURl
                                                End If
                                Next
                                
                                If blnClickLink = False Then
                                	UpdateReport "TESTSTEP", "", srtDescription, StepStartTime, Time, "Failed to click on <b>'" & strActionToBePerformed & "'</b> link", "Done"
                                    fn_ValidateEmailInOutlook = False
                                End If
                
                Case "VALIDATE"
                
                                If Instr(strBodyContent,strValue) > 0 Then
                                                UpdateReport "TESTSTEP", "",srtDescription, StepStartTime, Time, "<font color=""blue""><b><i>'" & strValue & "'</i></b></font> is displayed as the User ID in the Outlook message", "Done"
                                                fn_ValidateEmailInOutlook = True
                                Else
                                				UpdateReport "TESTSTEP", "",srtDescription, StepStartTime, Time, "'" & strValue & "' is not displayed as the User ID in the Outlook message", "Fail"
                                                fn_ValidateEmailInOutlook = False
                        						Environment.Value("TestStepLog") = "False"
												Environment.Value("TestObjectFlag") = "False"
                                End If
                                
                Case "VALIDATE EMAILADDRESS"
					
					            If Instr(strBodyContent,strValue) > 0 Then
					                            UpdateReport "TESTSTEP", "",srtDescription, StepStartTime, Time, "<font color=""blue""><b><i>'" & strValue & "'</i></b></font> is displayed in the Outlook message", "Done"
					                            fn_ValidateEmailInOutlook = True
					            Else
					            				UpdateReport "TESTSTEP", "",srtDescription, StepStartTime, Time, "'" & strValue & "' is not displayed as the Email Address in the Outlook message", "Fail"
					                            fn_ValidateEmailInOutlook = False
					    						Environment.Value("TestStepLog") = "False"
												Environment.Value("TestObjectFlag") = "False"
					            End If
                
End Select


End Function
'==================================================================================================================================================
' Name of the Function                    : fn_ReadEmailFromOutlook
' Description                                         : This Function to open outlook, access inbox folder and read e-mail from inbox
' Date and / or Version                     : 6/10/2013
' Example Call                                       : fn_ReadEmailFromOutlook(strSenderdetails,strSubject)
'==================================================================================================================================================

Function fn_ReadEmailFromOutlook(strSenderdetails, strSubject,srtDescription)                                           'Function to InvokeOutlook and filter the email

Set olapp=createobject("outlook.application")
Set objNameSpace=olapp.getnamespace("mapi")
Set inbox=objNameSpace.getdefaultfolder(6)
Set mail=inbox.items
StepStartTime = Time
blnMailFound = False

intMailWaitTime = DataTable.Value("GParam_intMailWaitTime","GLOBALPARAMETERS")
For intMailWait = 0 to intMailWaitTime
	For each eml in mail
	                If eml.unread=True and instr(Ucase(eml.subject), Ucase(strSubject))>0 then
	                                If Instr(ucase(eml.SenderName), Ucase(strSenderdetails))>0 and Instr(eml.creationtime,Date)>0 Then
	                                                strbodycontent=Rtrim(eml.body)
	                                                fn_ReadEmailFromOutlook = strbodycontent
	                                                blnMailFound = True
	                                                eml.unread = False
	                                                Exit For
	                                End If
	                End if
	Next
	If blnMailFound Then
		Exit For
	ElseIf (intMailWait = intMailWaitTime) and  blnMailFound = False Then
		UpdateReport "TESTSTEP", "",srtDescription, StepStartTime, Time, "Email not recevived from Sender " & strSenderdetails & " with subject " & strSubject, "Fail"
	End If
Wait(1)
Next

set olapp=nothing
set objnamespace=nothing
set inbox=nothing

End Function

'==================================================================================================================================================
' Name of the Function     : fn_Sendkeys
' Description                           :  This function is for using SendKeys in Application
' Date and / or Version       : 
' Example Call                         : fn_Sendkeys("","%FO")
'==================================================================================================================================================

Function fn_Sendkeys(strScenarioName,strDataVal)

					arrDataVal = Split(strDataVal,",")
					strKeyOperation = arrDataVal(0)
					
                    StepStartTime = Time
                    blnObjStatus = False
                    
                    strObjectHierarchy = Datatable.value("APP_SCREEN_NAME",strScenarioName)                              ' Object Hierarchy value ex: Browser("Login").Page("Login")
                    strObject = Datatable.Value("OBJECT", strScenarioName)
                    Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
                    
                    If strObject <> ""  Then
                       strActualObject = strObjectHierarchy & "." & strObject
                    ElseIf strObject = "" Then
                       strActualObject = strObjectHierarchy
                    End If
                    
                    Set ActualObject = Eval(strActualObject)
                    blnObjStatus = fnc_wait(ActualObject)
                    
                    If blnObjStatus Then
                    	Wait(3)
	            		Set WshShell = CreateObject("WScript.Shell")    
	                    WshShell.SendKeys strKeyOperation                                                                                                                                                                                        
	                    UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Operation performed Successfully", "Done" 
	                    Set WshShell = nothing
	                    fn_Sendkeys = True
	                Else
	                	UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Failed to perform the Operation", "Fail"
	                    fn_Sendkeys = False
                    End If
                   
                    
End Function

'==================================================================================================================================================
' Name of the Function     			  : OperateOnWinList
' Description       		   		 	     : This function is used to select or retrieve the values from the WinList object
' Date and / or Version       	    : 
' Example Call							 : OperateOnWinList(strObjectHierarchy, strObjectName, strAction_Name, strData)
'==================================================================================================================================================
Function OperateOnWinList(strObjectHierarchy, strObjectName, strAction_Name, strData)

		If  strAction_Name <> "STOREVALUE" and Instr(strData, "VAR_") = 1 Then				' If the value to be taken from already saved variable
				strData = Trim(Environment.Value(strData))
		End If

	    If  Instr(strData, "fn_") = 1 Then
			strData = Environment.Value(strData)
		End If
		
		strResData = Replace(strData,"/","")
		strResData = Replace(strResData,"-","")
		
	    Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.WinList(strObjectName)
		
		StepStartTime = Time
        If fnc_wait(ActualObject) = "True" Then
				Select Case strAction_Name
					  Case "SELECTVALUE"
							  On Error Resume Next
							  ObjectHierarchy.WinList(strObjectName).Select strData
							  strData = Trim(strData)
							  strAppData = ObjectHierarchy.WinList(strObjectName).GetROProperty("Selection")							  

							  If StrComp(Trim(strAppData),strData) = 0 Then
							  		ObjectHierarchy.WinList(strObjectName).Activate strAppData
									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value: <font color=""blue""><b><i>" & strAppData & "</i></b></font> is selected from the Dropdown <b>"  &Mid(strObjectName,4,Len(strObjectName)-3)& "</b>", "Done"
							  elseIf StrComp(strAppData,strData) <> 0 Then
  									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value: <font color=""blue""><b><i></font>" & strAppData & "</i></b></font> is not available in the Dropdown <b>"  &Mid(strObjectName,4,Len(strObjectName)-3) &"</b>", "Fail"
									Environment.Value("TestStepLog") = "False"
									Environment.Value("TestObjectFlag") = "False"
							  End If
							  
					  Case "COMPAREVALUE"
						  If instr(strData,"fn_") >0 or Instr(strData,"VAR_") >0  Then
								strExp = Environment.Value(strData)
						  Else
								strExp = strData
						  End If
								
							  ObjectHierarchy.WinList(strObjectName).Select strData								
							  strActual = Trim(ObjectHierarchy.WinList(strObjectName).GetROProperty("Selection"))
							  If StrComp(strExp,strActual) = 0 Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Expected Value - "& strExp &" and Actual value - "& strActual & " are matching</font>", "Pass"
							  Else
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Value mismatch, Actual value is: <i>" & strActual & "</i></font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
							  
					  Case "CHECKEXIST"
							  strExp = strData
							  strAct = ObjectHierarchy.WinList(strObjectName).Exist(5)
							  If UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object exists as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object does not exist</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  ElseIf UCASE(strExp) = UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object does not exist as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strAct)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object exists</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
					 Case "COMPAREANDSELECT"
							  strDataVal = fn_ListDataEncapsulation(strScenarioName,strDataVal)
							  
								If strDataVal <> "" then
									  	  ObjectHierarchy.WinList(strObjectName).Select strDataVal
										  strAppData = ObjectHierarchy.WinList(strObjectName).GetROProperty("Selection")							  
								
										  If StrComp(Trim(strAppData),Trim(strDataVal)) = 0 Then
										  		ObjectHierarchy.WinList(strObjectName).Activate strAppData
												UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value: <b><i>" & strAppData & "</i></b> is selected from the Dropdown", "Done"
										  elseIf StrComp(strAppData,strData) <> 0 Then
												UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value: <b><i>" & strAppData & "</i></b> is not available in the Dropdown", "Fail"
												Environment.Value("TestStepLog") = "False"
												Environment.Value("TestObjectFlag") = "False"
										  End If
								else
												UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value: <b><i>" & strDataVal & "</i></b> is not available in the Dropdown", "Fail"
												Environment.Value("TestStepLog") = "False"
												Environment.Value("TestObjectFlag") = "False"
								End if
								
					 Case "CHECKVALUENOTPRESENTINLIST"
					 
							If instr(strData,"fn_") >0 or Instr(strData,"VAR_") >0  Then
								strExp = Environment.Value(strData)
							Else
								strExp = strData
							End If
													
							strAllItems = Trim(ObjectHierarchy.WinList(strObjectName).GetROProperty("all items"))
							If Instr(strAllItems,strExp) = 0 Then
								UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">"& strExp &" not present in the list </font>", "Pass"
							Else
								UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red""> " & strExp & " present in the list </font>", "Fail"
								Environment.Value("TestStepLog") = "False"
							End If
							
					Case "CHECKVALUEPRESENTINLIST"
						
							If instr(strData,"fn_") >0 or Instr(strData,"VAR_") >0  Then
								strExp = Environment.Value(strData)
							Else
								strExp = strData
							End If
													
							strAllItems = Trim(ObjectHierarchy.WinList(strObjectName).GetROProperty("all items"))
							If Instr(strAllItems,strExp) > 1 Then
								UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">"& strExp &" present in the list </font>", "Pass"
							Else
								UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red""> " & strExp & " not present in the list </font>", "Fail"
								Environment.Value("TestStepLog") = "False"
							End If
							
				End Select
	   Else
                UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The WinList - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If
		
End Function

'==================================================================================================================================================
' Name of the Function     : fn_MenuSelection
' Description                           :  This function is for using SendKeys in Application
' Date and / or Version       : 
' Author			         : 	Abdul
' Example Call                         : fn_MenuSelection("View","Text")
'==================================================================================================================================================

Function fn_MenuSelection(strScenarioName,strDataVal)
					
			arrDataVal = Split(strDataVal,",")					
			strMenuName = arrDataVal(0) 
			strSelection = arrDataVal(1)

            StepStartTime = Time
            blnObjStatus = False
            
            strObjectHierarchy = Datatable.value("APP_SCREEN_NAME",strScenarioName)                              ' Object Hierarchy value ex: Browser("Login").Page("Login")
            strObject = Datatable.Value("OBJECT", strScenarioName)
            Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
            
            If strObject <> ""  Then
               strActualObject = strObjectHierarchy & "." & strObject
               Set ObjectHierarchy = Eval(strObjectHierarchy)
            ElseIf strObject = "" Then
               strActualObject = strObjectHierarchy
               Set ObjectHierarchy = Eval(strObjectHierarchy)
            End If
            
            Set ActualObject = Eval(strActualObject)
            blnObjStatus = fnc_wait(ActualObject)
            
            If blnObjStatus Then
        		Set WshShell = CreateObject("WScript.Shell")
        		ObjectHierarchy.Activate
				strMenu = Left(strMenuName,1)
				strItemToSelect = Left(strSelection,1)
                WshShell.SendKeys "%" & strMenu
                WshShell.SendKeys strItemToSelect
                UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Operation performed Successfully", "Done" 
                Set WshShell = nothing
                fn_MenuSelection = True
            Else
            	UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Failed to perform the Operation", "Fail"
                fn_MenuSelection = False
            End If
                   
End Function


'==================================================================================================================================================
' Name of the Function     : fn_ExtractListItem
' Description              :  This function is used for selecting list item from a Virtual list box object in the Application
' Created By			   : Srivaths
' Date and / or Version    : 
' Example Call             : fn_ExtractListItem("lstNextResOn","Text")
'==================================================================================================================================================

Function fn_ExtractListItem(strListBox,strListItem)

	On Error Resume Next

	blnListItemStatus = False
	
	DataTable.AddSheet("VirtualListItems")
	DataTable.ImportSheet Environment.Value("ExcelPath") & "VirtualListItems.xls",strListBox,"VirtualListItems"
	intReqListItem = DataTable.GetSheet("VirtualListItems").GetRowCount
	DataTable.GetSheet("VirtualListItems").SetCurrentRow(1)
	
	For intVirtualObjCounter = 1 to intReqListItem - 1
	
		strReqListItem = DataTable.Value("List_Items","VirtualListItems")
		If StrComp(strReqListItem,strListItem) = 0 Then
			blnExecutionStatus = True
			DataTable.GetSheet("VirtualListItems").SetCurrentRow(1)
			intVisibleVirtualObjCounter = 0
			Do
				strReqListItem = DataTable.Value("Visible_List_Items","VirtualListItems")
				intVisibleVirtualObjCounter = intVisibleVirtualObjCounter + 1
				DataTable.GetSheet("VirtualListItems").SetNextRow
			Loop Until(strReqListItem = "")
			intVisibleVirtualObjCounter = intVisibleVirtualObjCounter - 1
			Exit For
		
		ElseIf strReqListItem = "" or strListItem = "" Then
			Exit For
		End If
		
		DataTable.GetSheet("VirtualListItems").SetNextRow
		
	Next
	
	If blnExecutionStatus Then
		If intVirtualObjCounter > intVisibleVirtualObjCounter then
			intClickCounter = intVirtualObjCounter - intVisibleVirtualObjCounter
			fn_ExtractListItem = intVirtualObjCounter & ";" & intVisibleVirtualObjCounter & ";" & intClickCounter
		ElseIf (intVirtualObjCounter < intVisibleVirtualObjCounter) or (intVirtualObjCounter = intVisibleVirtualObjCounter) then
			intClickCounter = 0
			fn_ExtractListItem = intVirtualObjCounter & ";" & intVisibleVirtualObjCounter & ";" & intClickCounter
		End If
	End If
	
	DataTable.DeleteSheet("VirtualListItems")
	
End Function

'==================================================================================================================================================
' Name of the Function     : fn_SelectListItem
' Description              :  This function is used for selecting list item from a Virtual list box object in the Application
' Created By			   : Srivaths
' Date and / or Version    : 
' Example Call             : fn_SelectListItem("Sample","Text")
'==================================================================================================================================================

Function fn_SelectListItem(strScenarioName,strDataVal)
	
	On Error Resume Next
	
	arrDataVal = Split(strDataVal,",")
	strListItem = arrDataVal(0)
	
	StepStartTime = Time
	wdwWindow = Datatable.Value("APP_SCREEN_NAME", strScenarioName)              
	strObject = Datatable.Value("OBJECT", strScenarioName)
	Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
	
	arrVirtualObject = Split(strObject,"(")
	arrVirtualListObj = Split(arrVirtualObject(1),")")
	strVirtualReqObject = Replace(arrVirtualListObj(0),"""","")
	strVirtualReqObject =Trim(strVirtualReqObject)
	
	intObjCounter = fn_ExtractListItem(strVirtualReqObject,strListItem)
	
	arrVirtualObjCounter = Split(intObjCounter,";")
	intVirtualObjCounter = arrVirtualObjCounter(0)
	intVisVirtualObjCounter = arrVirtualObjCounter(1)
	intClickCounter = arrVirtualObjCounter(2)
	
	Set ObjectHierarchy = Eval(wdwWindow)
	Set ActualObject = ObjectHierarchy.VirtualList(strVirtualReqObject)
	
	If intClickCounter <> 0 Then
		For intClickOn = 1 to intClickCounter
			ObjectHierarchy.WinObject("btnScrollBar").VScroll micLineNext,1
		Next
	End If
	
	intVirtualObjCounter = intVirtualObjCounter - intClickCounter - 1
	intVirtualObjCounter = Cint(intVirtualObjCounter)
	
	ObjectHierarchy.Activate
	
	ActualObject.Select intVirtualObjCounter
	
	UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"<font color=""blue""><b><i>" & strDataVal & "</i></b></font> selected from the list successfully", "Done" 
	
End Function

'==================================================================================================================================================
' Name of the Function     : fn_ListDataEncapsulation
' Description              :  Makes user to select list objects by giving part names
' Created By                : Srivaths              
' Date and / or Version    : 
' Example Call             : fn_ListDataEncapsulation("Sample","Text")
'==================================================================================================================================================

Function fn_ListDataEncapsulation(strScenarioName,strDataVal)
	
		arrDataVal = Split(strDataVal,",")
		strItem = arrDataVal(0)
		If ubound(arrDataVal)>1 Then
			bPartString=True
			strItem1 = arrDataVal(1)
		End If
'		strProperty = arrDataVal(1)
		
		StepStartTime = Time
		strObjectHierarchy = Datatable.Value("APP_SCREEN_NAME", strScenarioName)              
		strObject = Datatable.Value("OBJECT", strScenarioName)
		
        If strObject <> ""  Then
           strActualObject = strObjectHierarchy & "." & strObject
           Set ObjectHierarchy = Eval(strObjectHierarchy)
        ElseIf strObject = "" Then
           strActualObject = strObjectHierarchy
           Set ObjectHierarchy = Eval(strObjectHierarchy)
        End If
        
        Set ActualObject = Eval(strActualObject)
        
        intActualObjectsCount = ActualObject.GetItemsCount
        For intCounter = 1 to intActualObjectsCount
        	strActualObjectValue = ActualObject.GetRoProperty("selection")
        	If Instr(strActualObjectValue,strItem) > 0 Then
        		If bPartString=true Then
        			If Instr(strActualObjectValue,strItem1) > 0 Then
	        		strItem = strActualObjectValue
	        		fn_ListDataEncapsulation = strItem
	        		Environment.Value("fn_ListDataEncapsulation") = strItem
	        		Exit For
	        	    End if
	        	else
	        	   	strItem = strActualObjectValue
	        		fn_ListDataEncapsulation = strItem
	        		Environment.Value("fn_ListDataEncapsulation") = strItem
	        		Exit For
        		End if
        		
        	End If
	    		ObjectHierarchy.Activate
				Wait(0.8)
				Set whShll = CreateObject("WScript.Shell")
				whShll.SendKeys "{DOWN} "
        Next
        
End Function

'==================================================================================================================================================
' Name of the Function     : fn_OperateOnWinobject
' Description              :  This function is used operate on shift keys
' Created By                                           : 
' Date and / or Version    : 
' Example Call             : fn_OperateOnWinobject("Sample","Text")
'==================================================================================================================================================
                                                                                                                                                
Function fn_OperateOnWinobject(strScenarioName,strData)
   
                StepStartTime = Time
                strObjectHierarchy = Datatable.Value("APP_SCREEN_NAME", strScenarioName)              
                strObject = Datatable.Value("OBJECT", strScenarioName)
                Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
                
                                Set ObjectHierarchy = Eval(strObjectHierarchy)
                                Set ActualObject = ObjectHierarchy.WinObject(strObjectName)
        						If fnc_wait(ActualObject) = "True" Then
                                                Wait(0.5)
                                                ObjectHierarchy.Activate
                                                ActualObject.Type strData
                                                If ActualObject Then
                                                        UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Button: <b>" & strData & "</b> is clicked", "Done"
                                                End If
                                End If

                                
End Function

'*******************************************************************************************************************************************
' Name: fnc_GetTimeStamp
' Author: Srivaths
' Description: Simple function that generates a time stamp for use in file names
' Inputs:	   none
' Output:	 string containing the timestamp
'*******************************************************************************************************************************************

Public Function fnc_GetTimeStamp()

	fnc_GetTimeStamp=  Date & "_" & Time
	fnc_GetTimeStamp= Replace(fnc_GetTimeStamp , "/", "-")
	fnc_GetTimeStamp= Replace(fnc_GetTimeStamp , ":", "_")

End Function

'==================================================================================================================================================
' Name of the Function     : fn_ValidateTableFieldValue
' Description              :  This function is used to validate data in a webtable
' Created By               : Srivaths
' Date and / or Version    : 22-Jul-13
' Example Call             : 
'==================================================================================================================================================

Function fn_ValidateTableFieldValue(strScenarioName,strDataVal)
                
    StepStartTime = Time
    blnValidation = False
    strNotPresentCounter = 1
    
    arrDataVal = Split(strDataVal,",")
    strExpectedFieldName = arrDataVal(0)
    strExpFieldValue = arrDataVal(1)
    
    'in case if the data to be verified is having "," then the expected value is passed using "#" instead of comma
    If instr(1,strExpFieldValue,"#") <> 0 Then
    	strExpFieldValue=replace(strExpFieldValue,"#",",")
    End If
    
    ObjectHierarchy = Datatable.Value("APP_SCREEN_NAME", strScenarioName)  
    strObject = Datatable.Value("OBJECT", strScenarioName)
    ActualObject = ObjectHierarchy & "." & strObject
    
    Set ObjectHierarchy = Eval(ObjectHierarchy)
    Set ActualObject = Eval(ActualObject)
    Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
    
    set objChildObjs=ActualObject.ChildObjects()
        
   For objIndex1=0 to objChildObjs.count-1
  
        if objChildObjs(objIndex1).getroproperty("micClass")="WebElement" then
        
            strFieldName=trim(objChildObjs(objIndex1).getroproperty("innertext"))
            
            if Trim(strFieldName)=Trim(strExpectedFieldName)  then
                strActulFldValue= objChildObjs(objIndex1+1).getroproperty("innertext")
                'Shafi - modified since actual string returns additional innertext
                If instr(1,trim(strActulFldValue),trim(strExpFieldValue)) Then
                	
                	If instr(1,strExpectedFieldName,":") <> 0  Then
                		strExpectedFieldName=replace(strExpectedFieldName,":","")
                	End If
                    bflg=true
                    Exit for
                End If
            End If
        End if
    Next
    
    If bflg=True Then
        UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value: <b><i>" & strExpFieldValue & "</i></b> is displayed as Field: <b><i>" & strExpectedFieldName & "</i></b>.", "Pass"
        'UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"The LPP Balance Due If terminated <font color=""blue""><b><i>'" & intAppBalanceDueIfTerminated & "'</i></b></font> displayed as expected", "Pass" 
    else
        UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value: <b><i>" & strAppData & "</i></b> is not displayed as Field: <b><i>" & strExpectedFieldName & "</i></b>.", "Fail"
        'UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Total Monthly LPP Amount and LPP Balance Due If terminated are not displayed as expected", "Fail" 
        Environment.Value("TestStepLog") = "False"
        Environment.Value("TestObjectFlag") = "False"
    End If
    
End Function

'==================================================================================================================================================
' Name of the Function     : fn_CaptureWebElement
' Description              : 
' Created By               : 
' Date and / or Version    : 22-Jul-13
' Example Call             : 
'==================================================================================================================================================

Function fn_CaptureWebElement(strScenarioName,strDataVal)
                
        StepStartTime = Time
        blnValidation = False
        ObjectHierarchy = Datatable.Value("APP_SCREEN_NAME", strScenarioName)  
        strObject = Datatable.Value("OBJECT", strScenarioName)
        ActualObject = ObjectHierarchy & "." & strObject
        
        Set ObjectHierarchy = Eval(ObjectHierarchy)
        Set ActualObject = Eval(ActualObject)
        Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
  		
		If ActualObject.Exist(10) then
			strInnerText = ActualObject.GetROProperty("innertext")
			Environment.Value(strDataVal)=strConfirmation

		End if
		
End function

'==================================================================================================================================================
' Name of the Function     			  : OperateOnPbEdit
' Description       		   		 	     : This function is used to enter or retrieve the values on / from the PbEdit object on Power Bulider Application
' Date and / or Version       	    :  12/11/2014, 
' Example Call							 : OperateOnPbButton("PbWindow("abc").PbButton(""efg"")","OK", "Set", "")
'==================================================================================================================================================
Function OperateOnPbEdit(strObjectHierarchy, strObjectName, strAction_Name, strData)
   
		If  strAction_Name <> "STOREVALUE" AND Instr(strData, "VAR_") = 1 Then				' If the value to be taken from already saved variable
			strData = Environment.Value(strData)
		End If

		If  Instr(strData, "fn_") = 1 Then
			strData = Environment.Value(strData)
		End If

		Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.PbEdit(strObjectName)
		StepStartTime = Time
        If fnc_wait(ActualObject) = "True" Then
				Select Case strAction_Name
					  Case "SETVALUE"	
	  						  On Error Resume Next
	  						  Wait(1)
	  						  arrDataVal = Split(strData,";")
	  						  If UBound(arrDataVal) > 1 Then
	  						  	strData = arrDataVal(2)
	  						  End If	  						  
	  						  ActualObject.Click
							  ActualObject.Set strData
							  strActualVal = ActualObject.GetROProperty("text")
							  
							  If StrComp(strActualVal,strData) = 0  Then					   		
							  		UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime, Time, "Value: <font color=""blue""> <b><i>" & strData & "</i></b></font> is entered " & " in <b>'" & Mid(strObjectName,4,Len(strObjectName)-3) & "'</b> field successfully", "Done"
							  ElseIf Err.Number <> 0 Then
							  		UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Value: <font color=""blue""><b><i>" & strDataVal & "</i></b></font> is not entered in field <b>" & strobjName &"</b> Due to Error" & Err.Description, "Fail"
							  Else
							  		UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Value: <font color=""blue""><b><i>" & strDataVal & "</i></b></font> is not entered in field <b>" & strobjName &"</b>", "Fail"
							  		Environment.Value("TestStepLog") = "False"
									Environment.Value("TestObjectFlag") = "False"
							  End If		
							  	
					   Case "SETPASSWORD"	
	  						  Wait(1)
	  							  						  
	  						  ActualObject.Click
							  ActualObject.SetSecure strData							  
							  UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Password is entered", "Done"
							  
					   Case "CLICK"
							  ActualObject.Click									
							  UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The WebEdit: <b>" & strObjectName & "</b> is clicked", "Done"
							
					  Case "TYPEVALUE"
							  Wait(1)
  							  ActualObject.Click
							  ActualObject.Type strData
							   If Instr(Environment.Value("strDescription"), "Password") > 0 Then
									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Password is entered", "Done"
							  Else
									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Value: <font color=""blue""><b><i>" & strData & "</i></b></font> is entered in field <b>" & Mid(strObjectName,4,Len(strObjectName)-3) &"</b> successfully", "Done"
							  End If
					  Case "COMPAREVALUE"
							   If Instr(strData,"fn_") >0 OR Instr(strData,"VAR_") >0  Then
									strExp = Environment.Value(strData)
							  Else
									strExp = strData
							  End If	
							  strActual = ObjectHierarchy.PbEdit(strObjectName).GetROProperty("text")
							  If Trim(UCASE(strExp)) = Trim(UCASE(strActual)) Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Expected Value - "& strExp &" and Actual value - "& strActual & " are matching</font>", "Pass"
							  Else
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Value mismatch, Expected Value is: <i>" & strExp & "</i>, and Actual value on the application is: <i>" & strActual & "</i></font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
					  Case "CHECKEXIST"
							  strExp = strData
							  strActual = ObjectHierarchy.PbEdit(strObjectName).Exist
							  If UCASE(strExp) = UCASE(CStr(strActual)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object exists as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strActual)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object does not exist</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  ElseIf UCASE(strExp) = UCASE(CStr(strActual)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object does not exist as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strActual)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object exists</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
'					 Case "SETPASSWORD"
'							 ActualObject.SetSecure strData
'							 UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Password is entered", "Done"
					 Case "STOREVALUE"
							 strVariableName = strData
							 Environment.Value("VAR_"& strVariableName) = ObjectHierarchy.PbEdit(strObjectName).GetROProperty("text")
							 'msgbox  Environment.Value("VAR_"& strVariableName)
							 UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The <i>"& strData &"</i> value: <i>"& Environment.Value("VAR_"& strVariableName)& "</i> is stored", "Done"
					 Case "CHECKENABLED"
								blnObjDisable= ObjectHierarchy.PbEdit(strObjectName).GetROProperty("enabled")								
								If blnObjDisable = "1" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Disabled as expected</font>", "Pass"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Enabled as expected</font>", "Pass"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Enabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								ElseIf blnObjDisable = "1" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Disabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								End If
						End Select
        Else
                UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The WebEdit - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If

End Function


'==================================================================================================================================================
' Name of the Function     			  : OperateOnPbEdit
' Description       		   		 	     : This function is used to enter or retrieve the values on / from the PbEdit object on Power Bulider Application
' Date and / or Version       	    :  12/11/2014, 
' Example Call							 : OperateOnPbButton("PbWindow("abc").PbButton(""efg"")","OK", "Set", "")
'==================================================================================================================================================
Function OperateOnWinEdit(strObjectHierarchy, strObjectName, strAction_Name, strData)
   
		If  strAction_Name <> "STOREVALUE" AND Instr(strData, "VAR_") = 1 Then				' If the value to be taken from already saved variable
			strData = Environment.Value(strData)
		End If

		If  Instr(strData, "fn_") = 1 Then
			strData = Environment.Value(strData)
		End If

		Set ObjectHierarchy = Eval(strObjectHierarchy)
		Set ActualObject = ObjectHierarchy.WinEdit(strObjectName)
		StepStartTime = Time
        If fnc_wait(ActualObject) = "True" Then
				Select Case strAction_Name
					  Case "SETVALUE"	
	  						  Wait(1)
	  						  
	  						  arrDataVal = Split(strData,";")
	  						  If UBound(arrDataVal) > 1 Then
	  						  	strData = arrDataVal(2)
	  						  End If	  						  
	  						  ActualObject.Click
							  ActualObject.Set strData
							  strActualVal = ActualObject.GetROProperty("text")
							  
							  If StrComp(strActualVal,strData) = 0  Then					   		
							  		UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime, Time, "Value: <font color=""blue""> <b><i>" & strData & "</i></b></font> is entered " & " in <b>'" & Mid(strObjectName,4,Len(strObjectName)-3) & "'</b> field successfully", "Done"
							  Else
							  		UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Value: <font color=""blue""><b><i>" & strDataVal & "</i></b></font> is not entered in field <b>" & strobjName &"</b>", "Fail"
							  		Environment.Value("TestStepLog") = "False"
									Environment.Value("TestObjectFlag") = "False"
							  End If		
							  	
					   Case "SETPASSWORD"	
	  						  Wait(1)
	  							  						  
	  						  ActualObject.Click
							  ActualObject.SetSecure strData							  
							  UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Password is entered", "Done"
							  
					   Case "CLICK"
							  ActualObject.Click									
							  UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The WebEdit: <b>" & strObjectName & "</b> is clicked", "Done"
							
					  Case "TYPEVALUE"
							  Wait(1)
  							  ActualObject.Click
							  ActualObject.Type strData
							   If Instr(Environment.Value("strDescription"), "Password") > 0 Then
									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Password is entered", "Done"
							  Else
									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Value: <font color=""blue""><b><i>" & strData & "</i></b></font> is entered in field <b>" & Mid(strObjectName,4,Len(strObjectName)-3) &"</b> successfully", "Done"
							  End If
					  Case "COMPAREVALUE"
							   If Instr(strData,"fn_") >0 OR Instr(strData,"VAR_") >0  Then
									strExp = Environment.Value(strData)
							  Else
									strExp = strData
							  End If	
							  strActual = ObjectHierarchy.WinEdit(strObjectName).GetROProperty("text")
							  If Trim(UCASE(strExp)) = Trim(UCASE(strActual)) Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Expected Value - "& strExp &" and Actual value - "& strActual & " are matching</font>", "Pass"
							  Else
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Value mismatch, Expected Value is: <i>" & strExp & "</i>, and Actual value on the application is: <i>" & strActual & "</i></font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
					  Case "CHECKEXIST"
							  strExp = strData
							  strActual = ObjectHierarchy.WinEdit(strObjectName).Exist
							  If UCASE(strExp) = UCASE(CStr(strActual)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object exists as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strActual)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object does not exist</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  ElseIf UCASE(strExp) = UCASE(CStr(strActual)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object does not exist as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strActual)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object exists</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
'					 Case "SETPASSWORD"
'							 ActualObject.SetSecure strData
'							 UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The Password is entered", "Done"
					 Case "STOREVALUE"
							 strVariableName = strData
							 Environment.Value("VAR_"& strVariableName) = ObjectHierarchy.WinEdit(strObjectName).GetROProperty("text")
							 'msgbox  Environment.Value("VAR_"& strVariableName)
							 UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The <i>"& strData &"</i> value: <i>"& Environment.Value("VAR_"& strVariableName)& "</i> is stored", "Done"
					 Case "CHECKENABLED"
								blnObjDisable= ObjectHierarchy.WinEdit(strObjectName).GetROProperty("enabled")								
								If blnObjDisable = "1" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Disabled as expected</font>", "Pass"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Field is Enabled as expected</font>", "Pass"
								ElseIf blnObjDisable = "0" AND UCASE(strData) = "FALSE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Enabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								ElseIf blnObjDisable = "1" AND UCASE(strData) = "TRUE" Then
										UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Field is Disabled</font>", "Fail"
										Environment.Value("TestStepLog") = "False"
								End If
					
					Case "FILESAVE"
							On Error Resume Next
								strFileName = strData
								Set objFolders = CreateObject("WScript.Shell").SpecialFolders
					 				MyDocumnetsFolder = objFolders("mydocuments")
						 			
						 		Set folderpath = CreateObject("Scripting.FileSystemObject")
						 			 Actualfolderpath = MyDocumnetsFolder &"\"& strFileName
					 			
					 			If folderpath.FileExists (Actualfolderpath) Then
									folderpath.DeleteFile(Actualfolderpath)
								End If
					 			
					 			ActualObject.Type Actualfolderpath
								
								If Err.Number = 0 Then
									UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime, Time, "The File <b><i>" & strFileName & "</i></b>  is saved in the Path <b><i>" & Actualfolderpath & "</i></b>", "Done"
								Else
									UpdateReport "TESTSTEP", "","<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The File <b>"&strFileName& "</b> is Not save in the Path <b>" & Actualfolderpath & "</b> Due to Error " & Err.Description & "</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
								End If
				
					Case "FILERESTORE"
			 				On Error Resume Next
							strFileName = strData
			 				Set objFolders = CreateObject("WScript.Shell").SpecialFolders
					 		 	MyDocumnetsFolder = objFolders("mydocuments")
						 		
						 	Set folderpath = CreateObject("Scripting.FileSystemObject")
						 	 	Actualfolderpath = MyDocumnetsFolder &"\"& strFileName
					 			
					 			ActualObject.Type Actualfolderpath
								
								If Err.Number = 0 Then
									UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime, Time, "The File <b><i>"& strFileName &"</i></b> is Retrieved from <b><i>"& Actualfolderpath & "</i></b> the Path", "Done"
								Else
									UpdateReport "TESTSTEP", "","<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The File <b>" & strFileName & "</b> is Not restored from the Path <b>" & Actualfolderpath & "</b> Due to Error " & Err.Description & "</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
								End If	
				End Select
        Else
                UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The WebEdit - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If

End Function

'==================================================================================================================================================
' Name of the Function     			  : OperateOnPbWindow
' Description       		   		 	     : This function is used to enter or retrieve the values on / from the PbEdit object on Power Bulider Application
' Author 							: Vivek Jain
'Date and / or Version       	    :  12/02/2015 
' Example Call							 : OperateOnPbWindow("PbWindow("abc")","OK", "Set", "")
'==================================================================================================================================================
Function OperateOnPbWindow(strObjectHierarchy, strObjectName, strAction_Name, strData)
   
		If  strAction_Name <> "STOREVALUE" AND Instr(strData, "VAR_") = 1 Then				' If the value to be taken from already saved variable
			strData = Environment.Value(strData)
		End If

		If  Instr(strData, "fn_") = 1 Then
			strData = Environment.Value(strData)
		End If

		Set ObjectHierarchy = Eval(strObjectHierarchy)
		StepStartTime = Time
        If fnc_wait(ActualObject) = "True" Then
				Select Case strAction_Name
					   Case "COMPAREVALUE"
							   If Instr(strData,"fn_") >0 OR Instr(strData,"VAR_") >0  Then
									strExp = Environment.Value(strData)
							  Else
									strExp = strData
							  End If	
							  strActual = ObjectHierarchy.GetROProperty("text")
							  If Trim(UCASE(strExp)) = Trim(UCASE(strActual)) Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Expected Value - "& strExp &" and Actual value - "& strActual & " are matching</font>", "Pass"
							  Else
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Value mismatch, Expected Value is: <i>" & strExp & "</i>, and Actual value on the application is: <i>" & strActual & "</i></font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
							  
					   Case "CHECKEXIST"
							  strExp = strData
							  strActual = ObjectHierarchy.Exist
							  If UCASE(strExp) = UCASE(CStr(strActual)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object exists as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strActual)) and UCASE(strExp) = "TRUE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object does not exist</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  ElseIf UCASE(strExp) = UCASE(CStr(strActual)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The object does not exist as expected</font>", "Pass"
							  ElseIf UCASE(strExp) <> UCASE(CStr(strActual)) and UCASE(strExp) = "FALSE" Then
									UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The object exists</font>", "Fail"
									Environment.Value("TestStepLog") = "False"
							  End If
'					 
					  Case "STOREVALUE"
							 strVariableName = strData
							 Environment.Value("VAR_"& strVariableName) = ObjectHierarchy.GetROProperty("text")
							 UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The <i>"& strData &"</i> value: <i>"& Environment.Value("VAR_"& strVariableName)& "</i> is stored", "Done"
					 
				End Select
        Else
                UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The WebEdit - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If

End Function