'==================================================================================================================================================
' Name of the Function     			  :fn_ValidataDataInPbDataWindowList
' Description       		   		  : This function is used  to Validate that a specific Value is not present pbDataWindow Grid
' Date and / or Version       	      : 03/03/2015
' Created By              			  : Anjan
' Example Call						  : Call fn_ValidataDataInPbDataWindowList("Test Scenario Name","URL") 
'==================================================================================================================================================               
Function fn_ValidataDataInPbDataWindowList(strScenarioName,strDataVal)
		
		StepStartTime = Time
		arrDataVal = Split(strDataVal,";")
		strRowVal = arrDataVal(0)
		strColumnVal = arrDataVal(1) 
        
        If Instr(1,arrDataVal(2),"VAR") <> 0 Then
        	strExpVal = Environment.Value(arrDataVal(1))
        Else
        	strExpVal = arrDataVal(2)
        End If
       
       ObjectHierarchy = Datatable.Value("APP_SCREEN_NAME", strScenarioName)  
        strObjectName = Datatable.Value("OBJECT", strScenarioName)
        ActualObject = ObjectHierarchy & "." & strObjectName
        
        Set ObjectHierarchy = Eval(ObjectHierarchy)
        Set ActualObject = Eval(ActualObject)
        Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
  		ObjectHierarchy.Activate
  		ActualObject.SelectCell strRowVal,strColumnVal
		blnflag = False
Do 
	ObjectHierarchy.Activate
	Set WshShell = CreateObject("WScript.Shell")    
	WshShell.SendKeys "{DOWN}"
	wait(2)
	strActual = ActualObject.GetCellData(strRowVal,strColumnVal)
	If strExpVal <> strActual  Then
		blflag = True
	End If
	If strTemp = strActual Then
		blnflag = True
		Exit Do
	End If
	strTemp = strActual
Loop
On Error Resume next
ActualObject.SetCellData strDataRowVal, strobjName,strDataVal
	If blflag AND Err.NUmber = "-10150" Then
		
		UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The Value -<b> "& strExpVal &"</b> is not Present as expected in the List object- <b><i>"& Mid(strObjectName,4,Len(strObjectName)-3) & "</i></b></font>", "Pass"
	Else
		UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Value -<b> "& strExpVal &"</b> is  Present in the List object - <b><i>"& Mid(strObjectName,4,Len(strObjectName)-3) & "</i></b></font>", "Fail"
	End If
End Function

'==================================================================================================================================================
' Name of the Function     			  :fn_VerifyDatainList
' Description       		   		  : This function is used  to Validate that a specific Value is not present pbDataWindow Grid
' Date and / or Version       	      : 03/03/2015
' Created By              			  : Anjan
' Example Call						  : Call fn_VerifyDatainList("Test Scenario Name","URL") 
'==================================================================================================================================================               
Function fn_VerifyDatainList(strScenarioName,strDataVal)
		
		StepStartTime = Time
		arrDataVal = Split(strDataVal,";")
		strRowVal = arrDataVal(0)
		strColumnVal = arrDataVal(1) 
        
        If Instr(1,arrDataVal(2),"VAR") <> 0 Then
        	strExpVal = Environment.Value(arrDataVal(1))
        Else
        	strExpVal = arrDataVal(2)
        End If
        
        On Error Resume Next
       	ObjectHierarchy = Datatable.Value("APP_SCREEN_NAME", strScenarioName)  
        strObject = Datatable.Value("OBJECT", strScenarioName)
        ActualObject = ObjectHierarchy & "." & strObject
        
        Set ObjectHierarchy = Eval(ObjectHierarchy)
        Set ActualObject = Eval(ActualObject)
        Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
  		ObjectHierarchy.Activate
  		ActualObject.SelectCell strRowVal,strColumnVal
		ActualObject.SetCellData strDataRowVal, strobjName,strDataVal
		
		If Err.NUmber = "-10150" Then
			UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">The Value -<b> "& strExpVal &"</b> is not Present as expected in the List object- <b><i>"& Mid(strObjectName,4,Len(strObjectName)-3) & "</i></b></font>", "Pass"
		Else
			UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Value -<b> "& strExpVal &"</b> is  Present in the List object - <b><i>"& Mid(strObjectName,4,Len(strObjectName)-3) & "</i></b></font>", "Fail"
		End If
End Function

'==================================================================================================================================================
' Name of the Function     			  :fn_VerifyDataDeletedFromPbDataWindowGrid
' Description       		   		  : This function is used  to VAlidate that a specific data is not present pbDataWindow Grid
' Date and / or Version       	      : 03/03/2015
' Created By              			  : Vivek
' Example Call						  : Call fn_VerifyDataDeletedFromPbDataWindowGrid("Test Scenario Name","URL") 
'==================================================================================================================================================               
                
Function fn_VerifyDataDeletedFromPbDataWindowGrid(strScenarioName,strDataVal)
		
		arrDataVal = Split(strDataVal,";")
		StepStartTime = Time
        strExpText = strDataVal
        
        ObjectHierarchy = Datatable.Value("APP_SCREEN_NAME", strScenarioName)  
        strObject = Datatable.Value("OBJECT", strScenarioName)
        ActualObject = ObjectHierarchy & "." & strObject
        
        Set ObjectHierarchy = Eval(ObjectHierarchy)
        Set ActualObject = Eval(ActualObject)
        Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
  		ObjectHierarchy.Activate  		
		RowCount = ActualObject.RowCount
		ColumnCount = ActualObject.ColumnCount
		
		If RowCount = 0 Then
		
		blnflag = false
		
		Else 
		For intRowCount = 1 To RowCount Step 1
				'ActualObject.SelectCell "#"&intRowCount, 1
				ObjectHierarchy.Activate 
				strActual = ActualObject.GetCellData("#"&intRowCount, "#1")
		
				If Instr(1, strActual, ".0") <> 0 Then
					strActual = cint(strActual)
				ElseIf Instr(1, strActual,"/") <> 0  Then			
					arrDataVal(0) = CDate(arrDataVal(0))		
					strActual = CDate(strActual)
				ElseIf Instr(1, strActual,"/") <> 0 AND len(strActual) > 10 Then
					strActual = CDate(Left(strActual,10))
					arrDataVal(0) = CDate(arrDataVal(0))								 							
				End If
				
			    If Trim(UCASE(arrDataVal(0))) = Trim(UCASE(strActual))  Then
					blnflag = True
					For intColumnCount = 2 To Ubound(arrDataVal)+1 Step 1
						ObjectHierarchy.Activate 
						strActualColumnData = ActualObject.GetCellData("#"&intRowCount, "#"&intColumnCount)
						arrIndex = intColumnCount-1
						
						If Instr(1, strActualColumnData, ".0") <> 0 Then 'This is to Check whethere string actual value has ".0". if contain it will be converted into integer data s
							strActualColumnData = cint(strstrActualColumnDataActual)
						ElseIf Instr(1, strActualColumnData,"/") <> 0  Then			
							arrDataVal(arrIndex) = CDate(arrDataVal(arrIndex))		
							strActualColumnData = CDate(strActualColumnData)
						ElseIf Instr(1, strActualColumnData,"/") <> 0 AND len(strActualColumnData) > 10 Then
							strActualColumnData = CDate(Left(strActualColumnData,10))
							arrDataVal(arrIndex) = CDate(arrDataVal(arrIndex))								 							
						End If
						
						If Trim(UCASE(arrDataVal(arrIndex))) = Trim(UCASE(strActualColumnData))  Then
							 strActual = strActual&";"&strActualColumnData
							 blnflag = True
						Else
							 blnflag = False
						End If
					Next
				End If
			If blnflag = True Then
'				strActual.SelecCell "#"&intRowCount,"#1"
				ActualObject.ActivateCell "#"&intRowCount,"#1"
				Environment.Value("VAR_RowNo") = intRowCount
				Exit For
			End If
		
		Next
        End If		
		If blnflag = false Then
				UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">"&strExpText&" is not present in Datawindow</font>", "Pass"
			 Else
				UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">"&strExpVal&" is present in Datawindow</font>", "Fail"
		Environment.Value("TestStepLog") = "False"
		Environment.Value("TestObjectFlag") = "False"
		End If	
							 
End Function                

'==================================================================================================================================================
' Name of the Function     : fn_VerifyTextBoxEnabled
' Description              :  This function is to verify text box is enabled
' Date and / or Version    : 
' Author                   :  Sunitha
' Example Call             : fn_VerifyTextBoxEnabled(strScenarioName, strDataVal)
'==================================================================================================================================================
Function fn_VerifyTextBoxEnabled(strScenarioName, strDataVal)
        StepStartTime = Time
        ObjectHierarchy = Datatable.Value("APP_SCREEN_NAME", strScenarioName)  
        strObject = Datatable.Value("OBJECT", strScenarioName)
        ActualObject = ObjectHierarchy & "." & strObject
        
        Set ObjectHierarchy = Eval(ObjectHierarchy)
        Set ActualObject = Eval(ActualObject)
        arrDataVal = Split(strDataVal ,";")
        strRowDataVal = arrDataVal(0)
        strobjName = arrDataVal(1)
        strValue = arrDataVal(2)
        
        ActualObject.SelectCell strRowDataVal,strobjName
        ActualObject.Type strValue
        
        strActualValue = ActualObject.GetCellData(strRowDataVal,strobjName)
        If strActualValue = strValue  Then
               UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Inactive Reason Text box is enabled", "Pass"
        Else
               UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Inactive Reason Text box is enabled", "Fail"
        	   Environment.Value("TestStepLog") = "False"
        End If
End Function

'==================================================================================================================================================
' Name of the Function     : fn_ClickOnContextMenu
' Description              :  This function is used to select a Context menu after clicking on PB Button in Daily Data window in MCS MPA
' Created By               : Anjan
' Date and / or Version    : 24-Mar-15
' Example Call             : fn_ClickOnContextMenu("Sample","Text")
'==================================================================================================================================================

Function fn_ClickOnContextMenu(strScenarioName,strData)
		On Error Resume Next
		 StepStartTime = Time
         strObjectHierarchy = Datatable.value("APP_SCREEN_NAME",strScenarioName)                              
         strObject = Datatable.Value("OBJECT", strScenarioName)
         Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
         
         Set strObjectHierarchy = Eval(strObjectHierarchy)
         
         	strObjectGet = Split(strObject, "(", -1)
			strObjectNew = Split(strObjectGet(1), ")")
			strObjectType = strObjectGet(0)																								' Object Type ex: WebEdit
			strObjectVal = Split(strObjectNew(0),"""")
			strObjectName = strObjectVal(1)
            
            arrData = Split(strData,";")
            stroperationData = arrData(0)
            strDataVal = arrData(1)
            
            strObjectHierarchy.Activate
			strObjectHierarchy.PbButton(strObjectName).Click
			Set wsh = Createobject("wscript.shell")
			Wait(1)
			wsh.SendKeys stroperationData
			If Err.Number =0 Then
				UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime, Time, "The context Menu <b><i>" & strDataVal &"</i></b> is selected successfully from the Object <font color=""blue""><b><i>" & strObjectName &"</i></b></font>", "Done"
			Else
				UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The Select operation not performed on Button - <b><i>" & strObjectName & "</i></b></font>", "Fail"							  
				Environment.Value("TestStepLog") = "False"
			End	If
End Function

'==================================================================================================================================================
' Name of the Function     			  : fn_KillWSSMProcess
' Description       		   		  : This function is used  to Kill the wssmProcess to launch MCS Applications from windows Task Manager
' Date and / or Version       	      : 12/03/2015
' Created By              			  : Anjan
' Example Call						  : Call fn_KillWSSMProcess("Test Scenario Name","URL") 
'==================================================================================================================================================
Function fn_KillWSSMProcess(strScenarioName,strDataVal)
	
		 Set oServ = GetObject("winmgmts:")
		 Set cProc = oServ.ExecQuery("Select * from Win32_Process")

		 For Each oProc In cProc
			If oProc.Name = "wssm.exe" Then
			   errReturnCode = oProc.Terminate()
			End If
		 Next
		Call fn_CloseAllPBApplications(strScenarioName,strDataVal)
End Function

'==================================================================================================================================================
' Name of the Function     			  : fn_CloseAllPBApplications
' Description       		   		  : This function is used  to Quits the Running MCS Power Builder Applications from windows Task Manager when Failuer is encountered during execution to continue with next script
' Date and / or Version       	      : 12/03/2015
' Created By              			  : Anjan
' Example Call						  : Call fn_CloseAllPBApplications("Test Scenario Name","URL") 
'==================================================================================================================================================
Function fn_CloseAllPBApplications(strScenarioName,strDataVal)
	
		 Set oServ = GetObject("winmgmts:")
		 Set cProc = oServ.ExecQuery("Select * from Win32_Process")
	
		For Each oProc In cProc
		'Add Case to below select case statement with the new process Name that you need to Terminate. 
		'NOTE: It is 'case sensitive
		AppName = oProc.Name 
			Select Case AppName
				   Case "mcs_gqa.exe" 'Kill MCS GQA app
			  			 errReturnCode = oProc.Terminate()
						
					Case "mcs_cha.exe" 'Kill MCS CHA app
			  			  errReturnCode = oProc.Terminate()
					
					Case "mcs_saa.exe" 'Kill MCS SAA app
			  			  errReturnCode = oProc.Terminate()	
				  	
				  	Case "mcs_sma.exe" 'Kill MCS SMA app
			  			  errReturnCode = oProc.Terminate()		  
					
					Case "mcs_psa.exe" 'Kill MCS PSA app
			  			  errReturnCode = oProc.Terminate()
			
					Case "mcs_mpa.exe" 'Kill MCS MPA app
			  			  errReturnCode = oProc.Terminate()
			
					Case "mcsquery.exe" 'Kill MCS QUERY app
			  			  errReturnCode = oProc.Terminate()
			
					Case "mcsinfo.exe" 'Kill MCS INFO app
			  			  errReturnCode = oProc.Terminate()
			End Select
		Next

End Function
'==================================================================================================================================================
' Name of the Function     			  : fn_CheckDataAvailabilityInPbDatawindowGrid
' Description       		   		  : This function is used  to VAlidate the Data in pbDataWindow Grid
' Date and / or Version       	      : 02/20/2015
' Created By              			  : AKshatha
' Example Call						  : Call fn_CheckDataAvailabilityInPbDatawindowGrid("Test Scenario Name","URL") 
'==================================================================================================================================================
Function fn_CheckDataAvailabilityInPbDatawindowGrid(strScenarioName,strDataVal)

         On Error Resume Next
         StepStartTime = Time
         strObjectHierarchy = Datatable.value("APP_SCREEN_NAME",strScenarioName)                              
         strObject = Datatable.Value("OBJECT", strScenarioName)
         Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
         
         ActualObject = strObjectHierarchy & "." & strObject        
         Set strObjectHierarchy = Eval(strObjectHierarchy)
         Set ActualObject = Eval(ActualObject)
         arrObj = Split(strObject,"(")
         strRepObj = Split(arrObj(1),")")
             
        RowCount = ActualObject.RowCount  
        	 If RowCount = 0 Then                                      'When data is not present in datawindow
                
                If strDataVal = "TRUE" Then                     'When Testcase expects data to be  present
                   UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Data is not populated in " & strRepObj(0) &"</font>", "Fail"
                   Environment.Value("TestStepLog") = "False"
                ElseIf strDataVal = "FALSE" Then                'When Testcase expects no rows to be retrieved
                   UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Data is not populated in " & strRepObj(0) &"</font>", "Pass"
                End If
                
             ElseIf RowCount > 0 Then                                 'When data is present in datawindow
                
                If strDataVal = "TRUE" Then                     'When Testcase expects data to be present
                 	UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Data is populated in " & strRepObj(0) &"</font>", "Pass"
                 ElseIf strDataVal = "FALSE" Then               'When Testcase expects no rows to be retrieved
                    UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Data is populated in " & strRepObj(0) &"</font>", "Fail"
                    Environment.Value("TestStepLog") = "False"
                 End If
                 
             End If
End Function

'==================================================================================================================================================
' Name of the Function     			  : fn_ListValueSelectInPbDatawindowGrid
' Description       		   		  : This function is used  to VAlidate the Data in pbDataWindow Grid
' Date and / or Version       	      : 02/20/2015
' Created By              			  : AKshatha
' Example Call						  : Call fn_ListValueSelectInPbDatawindowGrid("Test Scenario Name","URL") 
'==================================================================================================================================================
Function fn_ListValueSelectInPbDatawindowGrid(strScenarioName,strDataVal)

       	 arrDataVal = Split(strDataVal,";")
       	 strColumnName = arrDataVal(0)
       	 strExpVal = arrDataVal(1)
     
         StepStartTime = Time
         strObjectHierarchy = Datatable.value("APP_SCREEN_NAME",strScenarioName)                              
         strObject = Datatable.Value("OBJECT", strScenarioName)
         Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
         
         ActualObject = strObjectHierarchy & "." & strObject        
         Set strObjectHierarchy = Eval(strObjectHierarchy)
         Set ActualObject = Eval(ActualObject)
         
         arrObject = Split(strObject,"(")
         arrObjectName = Split(arrObject(1),")")
         strReportObject = arrObjectName(0)
             
         RowCount = ActualObject.RowCount  
         If RowCount > 0 Then  
			
			For intRow = 1 To RowCount Step 1
				ActualObject.selectCell "#"&intRow,strColumnName
				If ActualObject.GetCellData("#"&intRow,strColumnName) = strExpVal Then
					blnflag = True
					Exit For
				Else
					ActualObject.selectCell "#"&intRow,strColumnName
				End If
					
			Next			
                
	        If blnflag Then
	    	  	UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value: <font color=""blue""><b><i>" & strExpVal & "</i></b></font> is selected from <b>"  & strReportObject & "</b>" & " in PbDataWindow", "Done"
			Else
				UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "The value: <font color=""red""> <b><i>" & trim(strExpVal) & " </i></b></font> is not available in the List <b>" & strReportObject & "</b>" , "Fail"
	        End If        
        ELse
			UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "There are no List Values: <font color=""red""> <b><i>" & trim(strDataVal) & " </i></b></font> in the <b>" & strReportObject & "</b>" , "Fail"
	    End IF          
End Function

'==================================================================================================================================================
' Name of the Function     			  : fn_LaunchMCS
' Description       		   		  : This function is used  to launch the Sempra MCS Applications
' Date and / or Version       	      : 01/12/2014
' Created By              			  : Anjan
' Example Call						  : Call fn_LaunchMCS("Test Scenario Name","URL") 
'==================================================================================================================================================
Function fn_LaunchMCS(strScenaraioName,strDataVal)
			
			StepStartTime = Time
			strScenarioName = Environment.Value("ScenarioName")
		    Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
			strData = Split(strDataVal,";")
			SystemUtil.Run strData(1),"",strData(0) 'Launch the Application from the specified Path
			Wait(2)			
			
			UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Application <b><i>" & strData(1) & "</i></b> is launched successfully", "Done"
End Function
				
'==================================================================================================================================================
' Name of the Function     : fn_ClearBTULowField
' Description              : This Function Clears the BTU field Value
' Created By               : Anjan
' Date and / or Version    : 01/21/2015
' Example Call             : Call fn_LaunchMCS("Test Scenario Name","Input Data")
'==================================================================================================================================================

Function fn_ClearBTULowField(strScenarioName,strDataVal)


        arrDataVal = Split(strDataVal ,";")
		strRowDataVal = arrDataVal(0)
		strobjName = arrDataVal(1)

		If  Instr(arrDataVal(0), "VAR_") = 1 Then				' If the value to be taken from already saved variable
			strDataRowVal = Environment.Value(arrDataVal(0))
		End If
							
        ObjectHierarchy = Datatable.Value("APP_SCREEN_NAME", strScenarioName)  
        strObject = Datatable.Value("OBJECT", strScenarioName)
        ActualObject = ObjectHierarchy & "." & strObject
        
        Set ObjectHierarchy = Eval(ObjectHierarchy)
        Set ActualObject = Eval(ActualObject)
        
		ObjectHierarchy.Activate  
		ActualObject.SelectCell strDataRowVal,strobjName 'Select the Cell in the PbDataWindow
		For intdel = 1 To 12 Step 1
			Set wclr = CreateObject("Wscript.Shell")  
			wclr.SendKeys("{DELETE}") 
		Next
		
End function

'==================================================================================================================================================
' Name of the Function     : fn_RightClick
' Description              :  This functioncan be used to right click in Application
' Date and / or Version    : 15/01/2014
' Example Call             : fn_RightClick("Sample","")
'Author                    : Anjan Kumar
'==================================================================================================================================================

Function fn_RightClick(strScenarioName,strDataVal)

					Wait(2)				
                    StepStartTime = Time
                    blnObjStatus = False
                    
                    wdwWindow = Datatable.value("APP_SCREEN_NAME",strScenarioName)                              ' Object Hierarchy value ex: Browser("Login").Page("Login")
                    strObject = Datatable.Value("OBJECT", strScenarioName)
                    Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
                    
                    
                    If strObject <> ""  Then
                       strActualObject = wdwWindow & "." & strObject
                    ElseIf strObject = "" Then
                       strActualObject = wdwWindow
                    End If
                    
                    Set ActualObject = Eval(strActualObject)
                    blnObjStatus = fnc_wait(ActualObject)
                    Set  strObjectHierarchy = Eval(wdwWindow)                 
                    
                    If blnObjStatus Then
                    
                    	strObjectHierarchy.Activate
						Set objMDR=CreateObject("Mercury.DeviceReplay") 'To get the screen co-ordinates 

						x = ActualObject.GetROProperty("abs_x") 
						y = ActualObject.GetROProperty("abs_y")
						height = ActualObject.GetROProperty("height")
						width = ActualObject.GetROProperty("width")
						objMDR.MouseClick x+(width/20),y+(height/2),2
						strObjectHierarchy.WinMenu("ContextMenu").Select "Find" 'Right Clicks on the screen and clicks on Find 
                    	UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time, "Operation performed Successfully", "Done" 
	            		
	                Else
	                	UpdateReport "TESTSTEP", "","<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>","<font color=""red"">" & "Failed to perform the Operation" & "</font>", "Fail"
	                    fn_RightClick = False
	                    Environment.Value("TestStepLog") = "False"
                    End If
                   
                    
End Function

'*******************************************************************************************************************************************************************************
' Name of the Function     			 : fn_CloseApplication
' Description       		   		 	      : This function is used to close the application
' Date and / or Version       	    : 02/26/2013
' Author									      : Shrinidhi
' Input Parameters					  : 
' Example Call							 : Call fn_CloseApplication("",""BankTest)

'*******************************************************************************************************************************************************************************
Function fn_CloseApplication(strScenarioName, strDataVal)
		StepStartTime = Time
		strObjectHierarchy =Datatable.Value("APP_SCREEN_NAME", strScenarioName)              
		strObject = Datatable.Value("OBJECT", strScenarioName)
		Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
	
		Set ActualObject = Eval(strObjectHierarchy)
		ActualObject.Activate
		ActualObject.Close
		Wait(5)

		If  Window("wndWindow").Exist(2) Then
			Window("wndWindow").WinButton("btnOK").Click
		End If
		
		UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime, Time, "The application is closed", "Done"

		' To check if the application is closd correctly
		'Set ActualObject = Eval(strObjectHierarchy)
		'Demo
'		If  ActualObject.Exist(3) Then
'				ActualObject.Activate		
'				ActualObject.Close
'		
'			If  Window("wndWindow").Exist(2) Then
'				Window("wndWindow").WinButton("btnOK").Click
'		
'			End If
'			
'		End If
'		Wait(5)

End Function

'*******************************************************************************************************************************************************************************
' Name of the Function     			 : fn_ClickOnDialogBox
' Description       		   		 	      : This function is used to click OK on the dialog box
' Date and / or Version       	    : 02/26/2013
' Author									      : Shrinidhi
' Input Parameters					  : 
' Example Call							 : Call fn_ClickOnDialogBox("","")

'*******************************************************************************************************************************************************************************
Function fn_ClickOnDialogBox(strScenarioName, strDataVal)
		StepStartTime = Time
		strObjectHierarchy =Datatable.Value("APP_SCREEN_NAME", strScenarioName)              
		strObject = Datatable.Value("OBJECT", strScenarioName)
		Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
		If strObject <> ""  Then
			strActualObject = strObjectHierarchy & "." & strObject
		ElseIf strObject = "" Then
			strActualObject = strObjectHierarchy
		End If
	
		Set ActualObject = Eval(strActualObject)
		If  ActualObject.Exist(2) Then
				ActualObject.Click
		End If

End Function

'*******************************************************************************************************************************************************************************
' Name of the Function     			 : fn_Save
' Description       		   		 	      : This function is used to Save the CSOOrder data in the application
' Date and / or Version       	    : 02/26/2013
' Author									      : Shrinidhi
' Input Parameters					  : 
' Example Call							 : Call fn_Save("",""BankTest)

'*******************************************************************************************************************************************************************************
Function fn_Save(strScenarioName, strDataVal)
		StepStartTime = Time
		strObjectHierarchy =Datatable.Value("APP_SCREEN_NAME", strScenarioName)
		Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
	
		Set ActualObject = Eval(strObjectHierarchy)
		Call fnc_Wait(ActualObject)
		Wait(1)
		ActualObject.Activate		' Activate the Search reult window
		Wait(0.2)
		
					On Error Resume Next
					strStorePath = strData
					arrpath = Split(strStorePath,"\")
					strAutomationFolder = arrpath(0)
					strFileName = arrpath(1)
					
					Set objFolder = CreateObject("WScript.Shell").SpecialFolders
		 				MyDocumnetsFolder = objFolders("mydocuments")
			 			
			 		Set folderpath = CreateObject("Scripting.FileSystemObject")
			 			 Actualfolderpath = MyDocumnetsFolder &"\"& strAutomationFolder
			 			If folderpath.FolderExists (Actualfolderpath) Then
							folderpath.DeleteFolder(Actualfolderpath)
							folderpath.CreateFolder(Actualfolderpath)
						Else
							folderpath.CreateFolder(Actualfolderpath)
						End If
			 			strCompletePath = Actualfolderpath &"\"& strFileName
			 			ActualObject.Type strCompletePath
					If Err.Number = 0 Then
						UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime, Time, "The File is saved", "Done"
					Else
						UpdateReport "TESTSTEP", "","<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The File <b>"&strFileName& "</b> is Not save in the Path <b>"&Actualfolderpath& "</b> Due to Error " & Err.Number & "</font>", "Fail"
					End If
							

End Function

'*******************************************************************************************************************************************************************************
' Name of the Function     			 : fn_DeleteFolder
' Description       		   		 : This function is used to Delete the Specified  folder from MYDocuments folder in Local drive
' Date and / or Version       	     : 03/30/2015
' Author							 : Anjan
' Input Parameters					 : 
' Example Call						 : Call fn_DeleteFolder("","C:\Users\UserID\Documents\MCS_Autoamtion")

'*******************************************************************************************************************************************************************************
Function fn_DeleteFolder(strScenarioName, strDataVal)
		 On Error Resume Next
		strStorePath = strDataVal
		arrpath = Split(strStorePath,"\")
		strAutomationFolder = arrpath(0)
		strFileName = arrpath(1)
								    
'		Set objFolders = CreateObject("WScript.Shell").SpecialFolders
'				MyDocumnetsFolder = objFolders("mydocuments")
 			MyFolder = Environment.Value("ResultPath")
 		Set folderpath = CreateObject("Scripting.FileSystemObject")
 			 Actualfolderpath = MyFolder & strAutomationFolder
			
		If folderpath.FolderExists (Actualfolderpath) Then
			folderpath.DeleteFolder(Actualfolderpath)
		End If
				
End Function

'*******************************************************************************************************************************************************************************
' Name of the Function     			 : fn_LaunchApp
' Description       		   		 : This function is used to Save the CSOOrder data in the application
' Date and / or Version       	     : 02/26/2013
' Author							 : Shrinidhi
' Input Parameters					 : 
' Example Call						 : Call fn_LaunchApp("","C:\Test")

'*******************************************************************************************************************************************************************************
Function fn_LaunchApp(strScenarioName, strDataVal)
		StepStartTime = Time
		Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
		SystemUtil.Run strDataVal
		UpdateReport "TESTSTEP", "",Environment.Value("strDescription"), StepStartTime, Time, "The CIS application is launched", "Done"

End Function

'*******************************************************************************************************************************************************************************
' Name of the Function     			 : fn_CombineDateTime
' Description       		   		 : This function is used to combine the date & time values
' Date and / or Version       	     : 02/26/2013
' Author							 : Shrinidhi
' Input Parameters					 : 
' Example Call						 : Call fn_CombineDateTime("","ABC,XYZ,OPQ")

'*******************************************************************************************************************************************************************************
Function fn_CombineDateTime(strScenarioName, strDataVal)
		StepStartTime = Time
		strDataSplit = Split(strDataVal,",")
		strDate = Environment.Value("VAR_"&strDataSplit(0))
		strTime = Environment.Value("VAR_"&strDataSplit(1))
		strVal = strDate & strTime
		strVar = "VAR_" & strDataSplit(2)
		Environment.Value(strVar) = strVal
End Function

'*******************************************************************************************************************************************************************************
' Name of the Function     			 : fn_CreateACNumber
' Description       		   		 	      : This function is used to create a 6 digit a/c number
' Date and / or Version       	    : 02/26/2013
' Author									      : Shrinidhi
' Input Parameters					  : None 
' Example Call							 : Call fn_CreateACNumber("","")

'*******************************************************************************************************************************************************************************
Function fn_CreateACNumber(strScenarioName, strDataVal)
		Environment.Value("fn_CreateACNumber") = "0" & Hour(Time) & Minute(Time) & Second(Time)		
End Function


'==================================================================================================================================================
' Name of the Function     : fn_ScheduledPayementsTableValidation
' Description              :  This function is used to operate on table in scheduled payments
' Created By               : Srivaths
' Date and / or Version    : 2-Jul-13
' Example Call             : fn_ScheduledPayementsTableValidation("Sample","Text")
'==================================================================================================================================================

Function fn_ScheduledPayementsTableValidation(strScenarioName,strDataVal)
                
        StepStartTime = Time
        blnObjPresentStatus = False
        blnValidation = False
        intLength = 0
        
        arrDataVal = Split(strDataVal,",")
        strPayDate = arrDataVal(0)
        strPaymentDetailsBankName = arrDataVal(1)
        strPaymentDetailsPaymentAmount = arrDataVal(2)
        strObjectType = arrDataVal(3)
        strValidation = arrDataVal(4)
        
        strUrl = Datatable.Value("GParam_MyAccountURL", "GLOBALPARAMETERS")  
        ObjectHierarchy = Datatable.Value("APP_SCREEN_NAME", strScenarioName)  
        strObject = Datatable.Value("OBJECT", strScenarioName)
        ActualObject = ObjectHierarchy & "." & strObject
        
        Set ObjectHierarchy = Eval(ObjectHierarchy)
        Set ActualObject = Eval(ActualObject)
        Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
        
        strInnerHtml = ActualObject.GetROProperty("innerhtml")
        arrInnerHtml = Split(strInnerHtml,"<TR class=table-alt-field") 'Split the innerhtml to get the cell data
        
        For intCellCounter = 1 To Ubound(arrInnerHtml)
            intCheckBoxCounter = intCellCounter - 1 
            intCellStart = Instr(arrInnerHtml(intCellCounter),"checkbox name=value(paymentIdCB" & intCheckBoxCounter & ")")
            intCellEnd = Instr(arrInnerHtml(intCellCounter),"View/Edit</A> </TD></TR>")
            intLength = intCellEnd - intCellStart
            If intLength <> 0 Then
                strReqText = Mid(arrInnerHtml(intCellCounter),intCellStart,intLength)
                If Instr(strReqText,strPayDate) > 0 Then 'Comparision block for Pay Date
	                If Instr(strReqText,strPaymentDetailsBankName) > 0 Then ' Comparision block for Payment Details
                        If Instr(strReqText,strPaymentDetailsPaymentAmount) > 0 Then ' Comparision block for Payment Amount
                            intReqObjRowNumber = 2 + intCellCounter ' Required Row Number which needs to be operated
                            blnObjPresentStatus = True
                            Exit For
                        End If
	                End If
                End If
            End If
        Next
        
        If (intLength = 0) or (blnObjPresentStatus = False) Then
            UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Failed to find Bank Name <b>'" & strPaymentDetailsBankName & "'</b> with Pay Date <font color=""blue""><b><i>'" & strPayDate &"'</i></b></font>", "Fail" 
            Environment.Value("TestStepLog") = "False"
            Exit Function
        End If
        
		Select Case Ucase(strObjectType)
		                
		Case "CHECKBOX"
		                
		    If StrComp(Ucase(strValidation),"CHECK") = 0 Then
		        ActualObject.ChildItem(intReqObjRowNumber,1,"WebCheckBox",0).Click
		        strCheckedStatus = ActualObject.ChildItem(intReqObjRowNumber,1,"WebCheckBox",0).GetRoProperty("Checked")
		        If strCheckedStatus = 1 Then
		            UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Bank Name <b>'" & strPaymentDetailsBankName & "'</b> with Pay Date <font color=""blue""><b><i>'" & strPayDate & "'</i></b></font> selected successfully", "Pass" 
		            blnValidation = True
		        else
		            UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Failed to select Bank Name <b>'" & strPaymentDetailsBankName & "'</b> with Pay Date <font color=""blue""><b><i>'" & strPayDate &"'</i></b></font>", "Fail" 
		            Environment.Value("TestStepLog") = "False"
		        End If
		    ElseIf StrComp(Ucase(strValidation),"UNCHECK") = 0 Then
		        strCheckedStatus = ActualObject.ChildItem(intReqObjRowNumber,1,"WebCheckBox",0).GetRoProperty("Checked")
		        If strCheckedStatus = 1 Then
		            ActualObject.ChildItem(intReqObjRowNumber,1,"WebCheckBox",0).Click
		        End If
		        
		        If strCheckedStatus = 0 Then
		            UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Bank Name <b>'" & strPaymentDetailsBankName & "'</b> with Pay Date <font color=""blue""><b><i>'" & strPayDate & "'</i></b></font> deselected successfully", "Pass" 
		            blnValidation = True
		        else
		            UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Failed to deselect Bank Name <b>'" & strPaymentDetailsBankName & "'</b> with Pay Date <font color=""blue""><b><i>'" & strPayDate &"'</i></b></font>", "Fail" 
		            Environment.Value("TestStepLog") = "False"
		        End If
		    End If
		
		Case "LINK"
		                
		    If StrComp(Ucase(strValidation),"VALIDATE") = 0 Then
		        If blnObjPresentStatus Then
		            UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Bank Name <b>'" & strPaymentDetailsBankName & "'</b> with Pay Date <font color=""blue""><b><i>'" & strPayDate & "'</i></b></font> is present in the table", "Pass" 
		            blnValidation = True
		        else
		            UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Bank Name <b>'" & strPaymentDetailsBankName & "'</b> with Pay Date <font color=""blue""><b><i>'" & strPayDate &"'</i></b></font> is not present in the table", "Fail" 
		            Environment.Value("TestStepLog") = "False"
		        End If
		    ElseIf StrComp(Ucase(strValidation),"CLICK") = 0 Then
		        Set oReqObject = Description.Create()
		        oReqObject("micclass").Value = "Link"
		        oReqObject("Index").Value = intReqObjRowNumber - 3
		        
		        ObjectHierarchy.Link(oReqObject).Click
		        Wait(1)
		        UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Clicked on View/Edit Link for Bank Name <b>'" & strPaymentDetailsBankName & "'</b> with Pay Date <font color=""blue""><b><i>'" & strPayDate & "'</i></b></font> successfully", "Pass" 
		        blnValidation = True
		    End If
		
		Case "WEBELEMENT"
		                
		    If StrComp(Ucase(strValidation),"VALIDATE") = 0 Then
		        If blnObjPresentStatus Then
		            UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Bank Name <b>'" & strPaymentDetailsBankName & "'</b> with Pay Date <font color=""blue""><b><i>'" & strPayDate & "'</i></b></font> is present in the table", "Pass" 
		            blnValidation = True
		        else
		            UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Bank Name <b>'" & strPaymentDetailsBankName & "'</b> with Pay Date <font color=""blue""><b><i>'" & strPayDate &"'</i></b></font> is not present in the table", "Fail" 
		            Environment.Value("TestStepLog") = "False"
		        End If
		    ElseIf StrComp(Ucase(strValidation),"CLICK") = 0 Then
		        ActualObject.ChildItem(intReqObjRowNumber,1,"WebElement",0).Click
		        UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Clicked on Bank Name <b>'" & strPaymentDetailsBankName & "'</b> with Pay Date <font color=""blue""><b><i>'" & strPayDate & "'</i></b></font> successfully", "Pass" 
		        blnValidation = True
		    End If
		
		End Select

        If (blnValidation = True) and (blnObjPresentStatus = True) Then
            fn_ScheduledPayementsTableValidation = True
        else
            fn_ScheduledPayementsTableValidation = False
        End If

End Function

'==================================================================================================================================================
' Name of the Function     : fn_ManageSchedulePaymentsTableValidation
' Description              :  This function is used to operate on table in Manage Schedule Payments
' Created By               : Srivaths
' Date and / or Version    : 22-Jul-13
' Example Call             : fn_ManageSchedulePaymentsTableValidation("Sample","Text")
'==================================================================================================================================================

Function fn_ManageSchedulePaymentsTableValidation(strScenarioName,strDataVal)
                
        StepStartTime = Time
        blnValidation = False
        strNotPresentCounter = 1
        
        arrDataVal = Split(strDataVal,",")
        strExpPayDate = arrDataVal(0)
        strExpPaymentAmount = arrDataVal(1)
        strExpBankAccount = arrDataVal(2)
        strValidation = arrDataVal(3)
        strClickOnLink = arrDataVal(4)
        
        
        ObjectHierarchy = Datatable.Value("APP_SCREEN_NAME", strScenarioName)  
        strObject = Datatable.Value("OBJECT", strScenarioName)
        ActualObject = ObjectHierarchy & "." & strObject
        
        Set ObjectHierarchy = Eval(ObjectHierarchy)
        Set ActualObject = Eval(ActualObject)
        Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
  		
  		Select Case Ucase(strValidation)
  	
  		Case "VALIDATE"
  				
			If ActualObject.Exist(10) then
							
				set objChildObjs=ActualObject.ChildObjects()
				
				For objIndex1=0 to objChildObjs.count-1
				
					if objChildObjs(objIndex1).getroproperty("micClass")="WebElement" then
					
						strActPaymentDate=trim(objChildObjs(objIndex1).getroproperty("innertext"))
							If Trim(strActPaymentDate)=trim(strExpPayDate) Then
								strActPaymentAmt=trim(objChildObjs(objIndex1+1).getroproperty("innertext"))
								If Trim(strActPaymentAmt)=Trim(strExpPaymentAmount ) Then
									strActBankAccount=trim(objChildObjs(objIndex1+2).getroproperty("innertext"))
									If Trim(strActBankAccount)=Trim(strExpBankAccount  ) Then
										bFlag=true
										Exit For
									End if
								End If
								
							End If
					End if
				Next
			End if	
			
			Select Case UCASE(strClickOnLink )
				Case "PRESENT"
					If bFlag=true Then
						UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Bank Name <font color=""blue""><b>'" & strActBankAccount   & "'</b></font> with Pay Date <font color=""blue""><b><i>'" & strActPaymentDate   & "'</i></b></font> and Payment Amount of <font color=""blue""><b><i>" & strActPaymentAmt   & "</i></b></font> is present in the table as expected", "Pass" 
				 		Environment.Value("TestStepLog") = "Pass"
				 	else
				 		UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Bank Name <font color=""blue""><b>'" & strActBankAccount   & "'</b></font> with Pay Date <font color=""blue""><b><i>'" & strActPaymentDate & "'</i></b></font> and Payment Amount of <font color=""blue""><b><i>" & strActPaymentAmt & "</i></b></font> is not present in the table", "Fail" 
				 		Environment.Value("TestStepLog") = "False"
					End If
					
				Case "NOT PRESENT"
					If bFlag=False Then
						UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Bank Name <font color=""blue""><b>'" & strExpBankAccount   & "'</b></font> with Pay Date <font color=""blue""><b><i>'" & strExpPayDate & "'</i></b></font> and Payment Amount of <font color=""blue""><b><i>" & strExpPaymentAmount & "</i></b></font> is not present in the table as expected", "Pass" 
				 		Environment.Value("TestStepLog") = "Pass"
				 	else
				 		UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Bank Name <font color=""blue""><b>'" & strExpBankAccount   & "'</b></font> with Pay Date <font color=""blue""><b><i>'" & strExpPayDate & "'</i></b></font> and Payment Amount of <font color=""blue""><b><i>" & strExpPaymentAmount & "</i></b></font> is present in the table", "Fail" 
				 		Environment.Value("TestStepLog") = "False"
					End If
			End Select
				
  		
  		Case "CLICKONLINK"
  			
  			
		''strSplitData = Split(strData,",")
		strAppBankName = Trim(strExpBankAccount)
		strAppPaymentDate = Trim(strExpPayDate)
		strEditCancel= Trim(strClickOnLink)
		
		
		set objChildObjs=ActualObject.ChildObjects()
		For objIndex1=0 to objChildObjs.count-1
		
			if objChildObjs(objIndex1).getroproperty("micClass")="Link" then
			
			
				Select Case strEditCancel
				Case "Edit"
					strAppBankNameAct=trim(objChildObjs(objIndex1-3).getroproperty("innertext"))
					strAppPaymentDateAct=trim(objChildObjs(objIndex1-5).getroproperty("innertext"))
					
					''Account Number
					if ucase(strAppBankName)=ucase(strAppBankNameAct) and ucase(strAppPaymentDate)=ucase(strAppPaymentDateAct) then
						If trim(objChildObjs(objIndex1).getroproperty("innertext"))= strEditCancel  Then
							bflg=true
							objChildObjs(objIndex1).click
							Exit for
						End If
					End if
				Case "Cancel"
					strAppBankNameAct=trim(objChildObjs(objIndex1-3).getroproperty("innertext"))
					strAppPaymentDateAct=trim(objChildObjs(objIndex1-5).getroproperty("innertext"))
							
					''Account Number
					if ucase(strAppBankName)=ucase(strAppBankNameAct) and ucase(strAppPaymentDate)=ucase(strAppPaymentDateAct) then
						If trim(objChildObjs(objIndex1+1).getroproperty("innertext"))=strEditCancel  Then
							bflg=true
							objChildObjs(objIndex1+1).click
							Exit for
						End If
					End if
				End Select
			End if
		next
   
	   if bflg=true then
	    blnValidation=true
	    UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Clicked on <font color=""blue""><b><i>" & strClickOn & " </i></b></font> link successfully", "Pass" 
	   else
	    blnValidation=false
	    UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Manage Schedule Payment", "Fail" 
	             Environment.Value("TestStepLog") = "False"
	   End if
	   
     
    End Select

  
		If blnValidation = True Then
			fn_ManageSchedulePaymentsTableValidation = True
		else
			fn_ManageSchedulePaymentsTableValidation = False
		End If
        
End Function
'==================================================================================================================================================
' Name of the Function     : fn_ManageBankInformationTableValidation
' Description              :  This function is used to operate on table in Manage Bank Information
' Created By               : Srivaths
' Date and / or Version    : 22-Jul-13
' Example Call             : fn_ManageBankInformationTableValidation("Sample","Text")
'==================================================================================================================================================

Function fn_ManageBankInformationTableValidation(strScenarioName,strDataVal)
                
        StepStartTime = Time
        blnValidation = False
        
        arrDataVal = Split(strDataVal,",")
        strBankAccountName = arrDataVal(0)
        strType = arrDataVal(1)
        strValidation = arrDataVal(2)
        strClickOnLink = arrDataVal(3)
        
        ObjectHierarchy = Datatable.Value("APP_SCREEN_NAME", strScenarioName)  
        strObject = Datatable.Value("OBJECT", strScenarioName)
        ActualObject = ObjectHierarchy & "." & strObject
        
        Set ObjectHierarchy = Eval(ObjectHierarchy)
        Set ActualObject = Eval(ActualObject)
        Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
  		
  		Select Case Ucase(strValidation)
  	
  		Case "VALIDATE"
  				
			If ActualObject.Exist(10) then
				intRowCount = ActualObject.RowCount
				
				For intRowCounter = 2 to intRowCount
						strAppBankAccountName = Trim(ActualObject.ChildItem(intRowCounter,1,"WebElement",0).GetRoProperty("innertext"))
						strAppType= Trim(ActualObject.ChildItem(intRowCounter,2,"WebElement",0).GetRoProperty("innertext"))
						
					   	If (strAppBankAccountName = strBankAccountName) and (strAppType = strType) Then
						 	UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Bank Account <b>'" & strBankAccountName & "'</b> with type <font color=""blue""><b><i>'" & strType & "'</i></b></font> is present in the table", "Pass" 
						    blnValidation = True
						    Exit For
					   	End If
				Next
				
				If blnValidation = False Then
					UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Bank Account <b>'" & strBankAccountName & "'</b> with type <font color=""blue""><b><i>'" & strType & "'</i></b></font> is not present in the table", "Fail" 
					Environment.Value("TestStepLog") = "False"
				End If
			
			else
				UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Manage Bank information table is not present", "Fail" 
				Environment.Value("TestStepLog") = "False"
			End If
  				
  		Case "CLICKONLINK"
  			
  			
  	''strSplitData = Split(strData,",")
   strBankAccountName = Trim(strBankAccountName)
   strType = Trim(strType)
   strViewDelete = Trim(strClickOnLink)
     
   
   set objChildObjs=ActualObject.ChildObjects()
   For objIndex1=0 to objChildObjs.count-1
  
    if objChildObjs(objIndex1).getroproperty("micClass")="Link" then
     
     
     Select Case strViewDelete
     	Case "View"
     		strActAcntName=trim(objChildObjs(objIndex1-3).getroproperty("innertext"))
		     strAcntType=trim(objChildObjs(objIndex1-2).getroproperty("innertext"))
		     
		     ''Account Number
		     if ucase(strActAcntName)=ucase(strBankAccountName) and ucase(strAcntType)=ucase(strType) then
			      If objChildObjs(objIndex1).getroproperty("innertext")=strViewDelete  Then
			       		bflg=true
			       		objChildObjs(objIndex1).click
		                Exit for
			      End If
		     End if
     	Case "Delete"
		     strActAcntName=trim(objChildObjs(objIndex1-3).getroproperty("innertext"))
		     strAcntType=trim(objChildObjs(objIndex1-2).getroproperty("innertext"))
		     
		     ''Account Number
		     if ucase(strActAcntName)=ucase(strBankAccountName) and ucase(strAcntType)=ucase(strType) then
			      If objChildObjs(objIndex1+1).getroproperty("innertext")=strViewDelete  Then
			       	bflg=true
			       	objChildObjs(objIndex1+1).click
	                Exit for
			      End If
		     End if
     End Select
     
     
      
     End if
     
   next
   
   if bflg=true then
    blnValidation=true
    UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Clicked on <font color=""blue""><b><i>" & strClickOn & " </i></b></font> link successfully", "Pass" 
   else
    blnValidation=false
    UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Multiple Addresses? Link Accounts table is not present", "Fail" 
             Environment.Value("TestStepLog") = "False"
   End if
   
     
    End Select

  
		If blnValidation = True Then
			fn_ManageBankInformationTableValidation = True
		else
			fn_ManageBankInformationTableValidation = False
		End If
        
End Function

'==================================================================================================================================================
' Name of the Function     : fn_MultipleAddressesTableValidation
' Description              :  This function is used to operate on table in Multiple Addresses Link Accounts
' Created By               : Srivaths
' Date and / or Version    : 22-Jul-13
' Example Call             : fn_MultipleAddressesTableValidation("Sample","Text")
'==================================================================================================================================================

Function fn_MultipleAddressesTableValidation(strScenarioName,strDataVal)
                
        StepStartTime = Time
        blnValidation = False
        
        arrDataVal = Split(strDataVal,",")
        strAccountName = arrDataVal(0)
        strAccountNumber = arrDataVal(1)
        strValidation = arrDataVal(2)
        strClickOnLink = arrDataVal(3)
        
        ObjectHierarchy = Datatable.Value("APP_SCREEN_NAME", strScenarioName)  
        strObject = Datatable.Value("OBJECT", strScenarioName)
        ActualObject = ObjectHierarchy & "." & strObject
        
        Set ObjectHierarchy = Eval(ObjectHierarchy)
        Set ActualObject = Eval(ActualObject)
        Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
  		
  		Select Case Ucase(strValidation)
  	
  		Case "VALIDATE"
  				
			If ActualObject.Exist(10) then
				intRowCount = ActualObject.RowCount
				
				For intRowCounter = 2 to intRowCount
						strAppAccountName = Trim(ActualObject.ChildItem(intRowCounter,1,"WebElement",0).GetRoProperty("innertext"))
						strAppAccountNumber= Trim(ActualObject.ChildItem(intRowCounter,2,"WebElement",0).GetRoProperty("innertext"))
						
					   	If (strAppAccountName = strAccountName) and (strAppAccountNumber = strAccountNumber) Then
						 	UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Account Name <b>'" & strAppAccountName & "'</b> with Account Number <font color=""blue""><b><i>'" & strAppAccountNumber & "'</i></b></font> is present in the table", "Pass" 
						    blnValidation = True
						    Exit For
					   	End If
				Next
				
				If blnValidation = False Then
					UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Account Name <b>'" & strAppAccountName & "'</b> with Number <font color=""blue""><b><i>'" & strAppAccountNumber & "'</i></b></font> is not present in the table", "Fail" 
					Environment.Value("TestStepLog") = "False"
				End If
			
			else
				UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Multiple Addresses? Link Accounts table is not present", "Fail"
				Environment.Value("TestStepLog") = "False"
			End If
  				
  		Case "CLICKONLINK"
  			
  			''strSplitData = Split(strData,",")
			strAccountNum = Trim(strAccountNumber)
			strAccountName = Trim(strAccountName)
			strRenameRemove = Trim(strClickOnLink)
  			
  			set objChildObjs=ActualObject.ChildObjects()

			For objIndex1=0 to objChildObjs.count-1
		
				if objChildObjs(objIndex1).getroproperty("micClass")="Link" then


					Select Case ucase(strRenameRemove)

						CASE "RENAME"
						       
                                			strAccNum=trim(objChildObjs(objIndex1-3).getroproperty("innertext"))
                                			strAccName=trim(objChildObjs(objIndex1-4).getroproperty("innertext"))
                             
                                     			If ucase(strAccName)=ucase(strAccountName)and strAccNum=strAccountNum  then
                                                      		If objChildObjs(objIndex1).getroproperty("innertext")=strRenameRemove  Then
                                                                          bflg=true
                                                                    		objChildObjs(objIndex1).click
				            					Exit for
								End If
						
							End if

						 CASE "REMOVE"
                                			strAccNum=trim(objChildObjs(objIndex1-3).getroproperty("innertext"))
                                			strAccName=trim(objChildObjs(objIndex1-4).getroproperty("innertext"))
                             
                                     			If ucase(strAccName)=ucase(strAccountName)and strAccNum=strAccountNum  then
                                                      		If objChildObjs(objIndex1+1).getroproperty("innertext")=strRenameRemove  Then
                                                                          bflg=true
                                                                    		objChildObjs(objIndex1+1).click
				            					Exit for
								End If
						
							End if

						 End Select
			
				End if
			Next

			
    if bflg=true then
    	blnValidation=true
    	UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Clicked on <font color=""blue""><b><i>" & strClickOn & " </i></b></font> link successfully", "Pass"
    else
    	blnValidation=false
    	UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Multiple Addresses? Link Accounts table is not present", "Fail"
             Environment.Value("TestStepLog") = "False"
   End if
  
		If blnValidation = True Then
			fn_MultipleAddressesTableValidation = True
		else
			fn_MultipleAddressesTableValidation = False
		End If
		End Select
		
        
End Function


'==================================================================================================================================================
' Name of the Function     : fn_AppointmentSelection
' Description              :  This function is used to select a appointment date for service orders
' Created By               : Srivaths
' Date and / or Version    : 26-Jul-13
' Example Call             : fn_AppointmentSelection("Sample","Text")
'==================================================================================================================================================

Function fn_AppointmentSelection(strScenarioName,strDataVal)
                
        StepStartTime = Time
        blnValidation = False
        blnFoundFlag = False
		intClickCounter = 1
		
        arrDataVal = Split(strDataVal,",")
        strActColDat = arrDataVal(0)
        strActRowData = arrDataVal(1)
        strValidation = arrDataVal(2)
        strRadioButtonSelection  = arrDataVal(3)
        
        ObjectHierarchy = Datatable.Value("APP_SCREEN_NAME", strScenarioName)  
        strObject = Datatable.Value("OBJECT", strScenarioName)
        ActualObject = ObjectHierarchy & "." & strObject
        
        Set ObjectHierarchy = Eval(ObjectHierarchy)
        Set ActualObject = Eval(ActualObject)
        Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
        
        strActColDat = Replace(strActColDat,"||",",") 
        intRowCount = ActualObject.RowCount
        intColCount = ActualObject.ColumnCount(1)
        
        Select Case Ucase(strRadioButtonSelection)
        
        Case "FREE APPOINTMENTS"
        
		  		Select Case Ucase(strValidation)
		  	
		  		Case "VALIDATE"
		  		
			        For intColCounter = 1 To intColCount
						strColmnData = ActualObject.GetCellData(1,intColCounter)
						If strComp(strColmnData,strActColDat) = 0 Then
							For intRowCounter = 2 To intRowCount
								strRowData = ActualObject.GetCellData(intRowCounter,intColCounter)
								If strComp(strRowData,strActRowData) = 0 Then
									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Appointment for Day : <b>'" & strActColDat & "'</b> on time :  <font color=""blue""><b><i>'" & strActRowData & "'</i></b></font> is available in the table for selection", "Pass" 
									blnFoundFlag = True
									Exit For
								End If
							Next
						End If
						
						If blnFoundFlag Then
							Exit For
						End If
			
						intClickCounter = intClickCounter + 1
						If intClickCounter = 5 Then
							ObjectHierarchy.WebElement("eleNextButton").Click
							Wait(1)
							intClickCounter = 1
						End If
			
					Next
					
					If blnFoundFlag = False Then
						UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Appointment for Day : <b>'" & strActColDat & "'</b> on time :  <font color=""blue""><b><i>'" & strActRowData & "'</i></b></font> is not available for selection", "Fail" 
						Environment.Value("TestStepLog") = "False"
					End If
				
				Case "CLICKONLINK"
				
					For intColCounter = 1 To intColCount
						strColmnData = ActualObject.GetCellData(1,intColCounter)
						If strComp(strColmnData,strActColDat) = 0 Then
							For intRowCounter = 2 To intRowCount
								strRowData = ActualObject.GetCellData(intRowCounter,intColCounter)
								If strComp(strRowData,strActRowData) = 0 Then
									ActualObject.ChildItem(intRowCounter,intColCounter,"WebElement",0).Click
									UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Free Appointment selected for Day : <b>'" & strActColDat & "'</b> on time :  <font color=""blue""><b><i>'" & strActRowData & "'</i></b></font>", "Pass" 
									blnFoundFlag = True
									Exit For
								End If
							Next
						End If
						
						If blnFoundFlag Then
							Exit For
						End If
					
						intClickCounter = intClickCounter + 1
						If intClickCounter = 5 Then
							ObjectHierarchy.WebElement("eleNextButton").Click
							Wait(1)
							intClickCounter = 1
						End If
					
					Next
					
					If blnFoundFlag = False Then
						UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Free Appointment for Day : <b>'" & strActColDat & "'</b> on time :  <font color=""blue""><b><i>'" & strActRowData & "'</i></b></font> is not available for selection", "Fail" 
						Environment.Value("TestStepLog") = "False"
					End If
					
				End Select
				
		Case "CUSTOM APPOINTMENTS"
			
			For intColCounter = 1 To intColCount
				strColmnData = ActualObject.GetCellData(1,intColCounter)
				If strComp(strColmnData,strActColDat) = 0 Then
					ActualObject.ChildItem(2,intColCounter,"WebList",0).Select strActRowData
					UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Custom Appointment selected for Day : <b>'" & strActColDat & "'</b> on time :  <font color=""blue""><b><i>'" & strActRowData & "'</i></b></font>", "Pass" 
					blnFoundFlag = True
				End If
				
				If blnFoundFlag Then
					Exit For
				End If
			
				intClickCounter = intClickCounter + 1
				If intClickCounter = 5 Then
					ObjectHierarchy.WebElement("eleNext").Click
					Wait(1)
					intClickCounter = 1
				End If
			
			Next
			
			If blnFoundFlag = False Then
				UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Appointment for Day : <b>'" & strActColDat & "'</b> on time :  <font color=""blue""><b><i>'" & strActRowData & "'</i></b></font> is not available for selection", "Fail" 
				Environment.Value("TestStepLog") = "False"
			End If
					
		End Select
		
		
		
		If blnValidation = True Then
			fn_AppointmentSelection = True
		else
			fn_AppointmentSelection = False
		End If
			
End Function

'==================================================================================================================================================
' Name of the Function     : fn_EnrollLPPMyAccountValidation
' Description              :  This function is used to select a appointment date for service orders
' Created By               : Srivaths
' Date and / or Version    : 26-Jul-13
' Example Call             : fn_EnrollLPPMyAccountValidation("Sample","Text")
'==================================================================================================================================================

Function fn_EnrollLPPMyAccountValidation(strScenarioName,strDataVal)
                
        StepStartTime = Time
        blnValidation = False
        blnFoundFlag = False
		intClickCounter = 1
		
      
        
        ObjectHierarchy = Datatable.Value("APP_SCREEN_NAME", strScenarioName)  
        
        Set ObjectHierarchy = Eval(ObjectHierarchy)
        Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)

		Environment.Value("strOutstandingBalance") = ObjectHierarchy.WebElement("eleOutstandingBalance").GetRoProperty("innertext")
  		Environment.Value("fn_EnrollLPPMyAccountValidation") = ObjectHierarchy.WebElement("eleMonthlyLevelPayPlanAmount").GetRoProperty("innertext")
  		
  		blnValidation=false
		If blnValidation = True Then
			fn_EnrollLPPMyAccountValidation = True
			
		else
			fn_EnrollLPPMyAccountValidation = False
			UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"Total Monthly LPP Amount and LPP Balance Due are not displayed as expected", "Fail" 
		End If
			
End Function

'==================================================================================================================================================
' Name of the Function     : fn_ValidateDialogBoxMessage
' Description              : This Function Validates the Message displaye din the warning DialogBox 
' Created By               : Anjan Kumar
' Date and / or Version    : 27-Jan-15
' Example Call             : Call fn_ValidateDialogBoxMessage("Test Scenario Name","message")
'==================================================================================================================================================

Function fn_ValidateDialogBoxMessage(strScenarioName,strDataVal)
                
        Set objRegExp = New Regexp
		objRegExp.IgnoreCase = True
		objRegExp.Global = True
		objRegExp.Pattern = "[)(?*"",\\<>&#~%{}+_'.@:\/!;0-9 ]+"

        StepStartTime = Time
        strExpText = strDataVal
        
        ObjectHierarchy = Datatable.Value("APP_SCREEN_NAME", strScenarioName)  
        strObject = Datatable.Value("OBJECT", strScenarioName)
        ActualObject = ObjectHierarchy & "." & strObject
        
        Set ObjectHierarchy = Eval(ObjectHierarchy)
        Set ActualObject = Eval(ActualObject)
        Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
  		  		
		If ActualObject.Exist(3) then
			strActualText = ActualObject.GetROProperty("text")
            strActualDialogTxt = objRegExp.Replace(strActualText,"")
            strExpText = objRegExp.Replace(strExpText,"")
			If StrComp(Trim(strActualDialogTxt),Trim(strExpText)) = 0 Then
				UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"The Meassage <font color=""blue""><b><i>'" & strActualText & "'</i></b></font> is displayed as expected in DialogBox", "Pass" 
		    	
	  		Else
	  			UpdateReport "TESTSTEP", "", Environment.Value("strDescription"), StepStartTime, Time,"The message is not displayed as expected in DialogBox", "Fail" 
            	Environment.Value("TestStepLog") = "False"	
			End If
		End If			
End function

'==================================================================================================================================================
' Name of the Function     : fn_CompareColumnsValueInPbDataWindow
' Description              : This Function Validates the all the columns Data in PbDataWindow List
' Created By               : Anjan Kumar
' Date and / or Version    : 02-Feb-15
' Example Call             : Call fn_CompareColumnsValueInPbDataWindow("Test Scenario Name","message")
'==================================================================================================================================================
					
Function fn_CompareColumnsValueInPbDataWindow(strScenarioName,strDataVal)
		
		arrDataVal = Split(strDataVal,";")
		StepStartTime = Time
        strExpText = strDataVal
        
        ObjectHierarchy = Datatable.Value("APP_SCREEN_NAME", strScenarioName)  
        strObject = Datatable.Value("OBJECT", strScenarioName)
        ActualObject = ObjectHierarchy & "." & strObject
        
        Set ObjectHierarchy = Eval(ObjectHierarchy)
        Set ActualObject = Eval(ActualObject)
        Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
  		ObjectHierarchy.Activate  		
		RowCount = ActualObject.RowCount
		ColumnCount = ActualObject.ColumnCount
		
		For intRowCount = 1 To RowCount Step 1
				'ActualObject.SelectCell "#"&intRowCount, 1
				ObjectHierarchy.Activate 
				strActual = ActualObject.GetCellData("#"&intRowCount, "#1")
		
				If Instr(1, strActual, ".0") <> 0 Then
					strActual = cint(strActual)
				ElseIf Instr(1, strActual,"/") <> 0  Then			
					arrDataVal(0) = CDate(arrDataVal(0))		
					strActual = CDate(strActual)
				ElseIf Instr(1, strActual,"/") <> 0 AND len(strActual) > 10 Then
					strActual = CDate(Left(strActual,10))
					arrDataVal(0) = CDate(arrDataVal(0))								 							
				End If
				
			    If Trim(UCASE(arrDataVal(0))) = Trim(UCASE(strActual))  Then
					blnflag = True
					For intColumnCount = 2 To Ubound(arrDataVal)+1 Step 1
						ObjectHierarchy.Activate 
						strActualColumnData = ActualObject.GetCellData("#"&intRowCount, "#"&intColumnCount)
						arrIndex = intColumnCount-1
						
						If Instr(1, strActualColumnData, ".0") <> 0 Then 'This is to Check whethere string actual value has ".0". if contain it will be converted into integer data s
							strActualColumnData = cint(strstrActualColumnDataActual)
						ElseIf Instr(1, strActualColumnData,"/") <> 0  Then			
							arrDataVal(arrIndex) = CDate(arrDataVal(arrIndex))		
							strActualColumnData = CDate(strActualColumnData)
						ElseIf Instr(1, strActualColumnData,"/") <> 0 AND len(strActualColumnData) > 10 Then
							strActualColumnData = CDate(Left(strActualColumnData,10))
							arrDataVal(arrIndex) = CDate(arrDataVal(arrIndex))								 							
						End If
						
						If Trim(UCASE(arrDataVal(arrIndex))) = Trim(UCASE(strActualColumnData))  Then
							 strActual = strActual&";"&strActualColumnData
							 blnflag = True
						Else
							 blnflag = False
						End If
					Next
				End If
			If blnflag = True Then
'				strActual.SelecCell "#"&intRowCount,"#1"
				ActualObject.ActivateCell "#"&intRowCount,"#1"
				Environment.Value("VAR_RowNo") = intRowCount
				Exit For
			End If
		
		Next							
		If blnflag Then
		UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Expected Value - "& strExpText &" and Actual value - "& strActual & " are matching</font>", "Pass"
		Else
		UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Value mismatch, Expected Value is: <i>" & strExpText & "</i>, and Actual value on the application is: <i>" & strActual & "</i></font>", "Fail"
		Environment.Value("TestStepLog") = "False"
		Environment.Value("TestObjectFlag") = "False"
		End If	
							 
End Function
							 
'==================================================================================================================================================
' Name of the Function     : fn_MenuSelection
' Description              :  This function is for using SendKeys in Application
' Date and / or Version    : 
' Author			       : 	Abdul
' Example Call             : fn_MenuSelection("View","Text")
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
' Name of the Function     : fn_VerifyGNNsDeleted
' Description              : Verifies the Added GNNs are Deleted from MCS Hourly Query Wimdow
' Created By               : Madhusudhana K S
' Date and / or Version    : 19-Feb-2015
' Example Call             : fn_VerifyGNNsDeleted()
'==================================================================================================================================================

Function fn_VerifyGNNsDeleted(strScenarioName,strDataVal)
                
        StepStartTime = Time
        blnValidation = False
        ObjectHierarchy = Datatable.Value("APP_SCREEN_NAME", strScenarioName)  
        strObject = Datatable.Value("OBJECT", strScenarioName)
        ActualObject = ObjectHierarchy & "." & strObject
        arrObj = Split(strObject,"(")
        strRepObj = Split(arrObj(1),")")     
        
        
        Set ObjectHierarchy = Eval(ObjectHierarchy)
        Set ActualObject = Eval(ActualObject)
        Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
  		Int Rowcount=0
		If ActualObject.Exist(10) then
			Rowcount = ActualObject.GetROProperty("RowCount")
			If Rowcount=0 Then
			UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">Data is not present in " & strRepObj(0) & "</font>", "Pass"
		Else
		UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">Data is present in " & strRepObj(0) & "</font>", "Fail"
		Environment.Value("TestStepLog") = "False"
		End If
		
		Else
				UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") & "</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">The RadioButton - <b><i>" & strObjectName & "</i></b> does not exist</font>", "Fail"
				Environment.Value("TestStepLog") = "False"
				Environment.Value("TestObjectFlag") = "False"
		End If
		
End function		
'==================================================================================================================================================
' Name of the Function     			  : fn_CheckSpecDataDoesNotExistInPbDatawindowGrid
' Description       		   		  : This function is used  to VAlidate that a specific data is not present pbDataWindow Grid
' Date and / or Version       	      : 02/27/2015
' Created By              			  : AKshatha
' Example Call						  : Call fn_CheckDataAvailabilityInPbDatawindowGrid("Test Scenario Name","URL") 
'==================================================================================================================================================
Function fn_CheckSpecDataDoesNotExistInPbDatawindowGrid(strScenarioName,strData)

         StepStartTime = Time
         strObjectHierarchy = Datatable.value("APP_SCREEN_NAME",strScenarioName)                              
         strObject = Datatable.Value("OBJECT", strScenarioName)
         Environment.Value("strDescription") = Datatable.Value("STEP_DESCRIPTION", strScenarioName)
         
         ActualObject = strObjectHierarchy & "." & strObject        
         Set strObjectHierarchy = Eval(strObjectHierarchy)
         Set ActualObject = Eval(ActualObject)
             
         arrDataVal = Split(strData,";")
	     strColumnName = arrDataVal(0)
	  	 strExpVal = arrDataVal(1) 
		  	
         RowCount = ActualObject.RowCount 
        
         If RowCount = 0 Then                                      'When data is not present in datawindow
                
			blnflag = True
			
         ElseIf RowCount > 0 Then                                 'When data is present in datawindow
                  
			If Instr(1,arrDataval(1),"VAR") = 1 Then ' TO use already stored value as expected value for comparision
               strExpVal = Environment.Value(arrDataval(1))
            End If 

		  	If Ubound(arrDataVal) = 3 Then
		  	    strSpecColumnName = arrDataVal(2)
		        strSpecRowVal = arrDataVal(3)
 			
	            For intRowCount = 1 To RowCount Step 1
							 	 
			     ActualObject.SelectCell "#"&intRowCount,strSpecColumnName
				 blnflag = True
				 If ActualObject.GetCellData("#"&intRowCount,strSpecColumnName) = strSpecRowVal Then
				 	strActualDataVal = ActualObject.GetCellData("#"&intRowCount,strColumnName)  
					If Instr(1,strActualDataVal, "/") <> 0 AND Instr(1,strExpVal, "/") <> 0  Then 'If the Actual Value is Date format and expected from data datasheet is date which treated as string is converted into CDate format
						 	
				        If StrComp(CDate(strActualDataVal),CDate(strExpVal)) = 0 Then
								blnflag = False
				        End If
						  
					ElseIF StrComp(Trim(strActualDataVal),Trim(strExpVal)) = 0 Then 'If the Actual Value and the expceted values are string 
					    		    blnflag = False
					End If 
					
			            If blnflag = False Then 'Exit the for Loop if comparision passed 
					 	   		Exit For						 		
					 	End If
			     End If
			  Next
			  
			Else
				  
				  For intRowCount = 1 To RowCount Step 1
				  
				     ActualObject.SelectCell "#"&intRowCount,strColumnName
				     blnflag = True
				   
				     strActualDataVal = ActualObject.GetCellData("#"&intRowCount,strColumnName)  
					 If Instr(1,strActualDataVal, "/") <> 0 AND Instr(1,strExpVal, "/") <> 0  Then 'If the Actual Value is Date format and expected from data datasheet is date which treated as string is converted into CDate format
						 	
				        If StrComp(CDate(strActualDataVal),CDate(strExpVal)) = 0 Then
								blnflag = False
				        End If
						  
					 ElseIF StrComp(Trim(strActualDataVal),Trim(strExpVal)) = 0 Then 'If the Actual Value and the expceted values are string 
					    		    blnflag = False
					 End If 
	
						If blnflag = False Then 'Exit the for Loop if comparision passed 
					 	   		Exit For						 		
					 	End If
				Next
		    End If
		End IF	
			
			If blnflag Then
				UpdateReport "TESTSTEP", "", "<font color=""green"">" & Environment.Value("strDescription") &"</font>", "<font color=""green"">" & StepStartTime & "</font>", "<font color=""green"">" & Time & "</font>", "<font color=""green"">"&strExpVal&" is not present in Datawindow</font>", "Pass"
			 Else
				UpdateReport "TESTSTEP", "", "<font color=""red"">" & Environment.Value("strDescription") &"</font>", "<font color=""red"">" & StepStartTime & "</font>", "<font color=""red"">" & Time & "</font>", "<font color=""red"">"&strExpVal&" is present in Datawindow</font>", "Fail"
					 
			 End If		 
End Function              
                