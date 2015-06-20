'Add reference: Xscript Actions
'Add reference: Xscript Classes
'Add reference: Xscript Caller
'Add reference: Xscript Excel

Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Globalization

Public Module InterpreterModule
		
	'Persistent objects
	Dim strActiveExeFileNameList As New List(Of String)
	
	'Local vars that need not be reset between instructions
	Dim intNumNothings As Integer=10
	Dim intTemp As Integer
	Dim strTemp As String
	Dim dblTemp As Double
	Dim blnTemp As Boolean
	Dim strTempArr() As String
		
'==============================================================================================================
' FnInterpret
'
' 		Translates Xscript commands and objects into meaningful values so they can be properly executed.
'
' 		@test:				The current test being executed.  This object already contains the current step.
'		@dtCurrentDateTime:	Variable that stores the datetime when Xscript was initialized.
'		@blnTrailingHyphen:	Boolean to indicate whether the user cares if the current step fails or not.
'		@dicConstants:		A dictionary of variables that help tweak behavior in the Xscript script
'==============================================================================================================

	Public Function FnInterpret(ByRef test As Test, ByRef dtCurrentDateTime As DateTime, ByRef blnTrailingHyphen As Boolean, ByRef dicConstants As Dictionary(Of String, String)) As String
	
		'Declare objects
		Dim objUserArgsList As New List(Of Object)
		
		'Declare variables
		Dim strCmd As String
		Dim blnInvertTest, blnOccasionalObjectNotPresent As Boolean
		Dim x, y As Integer
				
		'Default return
		FnInterpret=Nothing
		strCmd=Nothing
		
		'Prase the instruction into useful pieces
		FnInterpret = FnParseInstruction(test.strInstruction, strCmd, objUserArgsList)
		
		'If there are parameters for this instruction
		If String.IsNullOrEmpty(FnInterpret) AndAlso objUserArgsList.Count > 0 Then
			
			'For each element in the param array, we need to translate it into a corresponding value if it's not raw text
			For intIndex As Integer = 0 To objUserArgsList.Count-1
				
				'If there's a leading exclamation mark, set a boolean to invert the test and strip the first character
				If String.Equals(Left(objUserArgsList.Item(intIndex), 1), "!") Then
					blnInvertTest = True
					objUserArgsList.Item(intIndex) = Mid(objUserArgsList.Item(intIndex), 2)
				End If
				
				'If this is an integer, convert it
				If Integer.TryParse(objUserArgsList.Item(intIndex), intTemp) Then
					objUserArgsList.Item(intIndex) = intTemp
					
				'If this is an double, convert it
				Else If Double.TryParse(objUserArgsList.Item(intIndex), dblTemp) Then
					objUserArgsList.Item(intIndex) = dblTemp
					
				'If the param reads 'True' without the quotes, convert it to a boolean and disregard
				Else If String.Compare(objUserArgsList.Item(intIndex), "True", True) = 0 
					objUserArgsList.Item(intIndex) = True
					
				'If the param reads 'False' without the quotes, convert it to a boolean and disregard
				Else If String.Compare(objUserArgsList.Item(intIndex), "False", True) = 0 Then
					objUserArgsList.Item(intIndex) = False
				
				'If this is a coordinate, convert it to a proper Point and move on
				Else If FnIsCoordinate(objUserArgsList.Item(intIndex), x, y) Then
					
					'Ensure this item isn't already a Point object that was inserted by Xscript itself because the object is defined by coordinates					
					If Not String.Equals(objUserArgsList.Item(intIndex).GetType.ToString, "SilkTest.Ntf.Point") Then
						objUserArgsList.Item(intIndex) = New Point(x, y)
					Else
						FnInterpret = "Xscript cannot take action at a coordinate on this object because the object itself is defined by a coordinate within another object; " & _
									  "there is no zero-based reference point.  The object is likely defined as such because Silk cannot outright recognize the object with which " & _
									  "you need to interact.  Please remove the coordinate reference from your instruction."
					End If
					
				'If this is a MouseButton enumeration reference, convert it
				Else If String.Compare(Left(objUserArgsList.Item(intIndex), 12), "MouseButton.", True)=0 Then
					objUserArgsList.Item(intIndex) = FnTranslateMouseButton(objUserArgsList.Item(intIndex), FnInterpret)
						
				'If this is a ClickType enumeration reference, convert it
				Else If String.Compare(Left(objUserArgsList.Item(intIndex), 10), "ClickType.", True)=0 Then
					objUserArgsList.Item(intIndex) = FnTranslateClickType(objUserArgsList.Item(intIndex), FnInterpret)
					
				'If this is a ModifierKeys enumeration reference, convert it
				Else If String.Compare(Left(objUserArgsList.Item(intIndex), 13), "ModifierKeys.", True)=0 Then
					objUserArgsList.Item(intIndex) = FnTranslateModifierKeys(objUserArgsList.Item(intIndex), FnInterpret)
					
				'If this parameter is present in test.strLocatorDic, translate it into a control name, then append the parent at the beginning.  Also, check for trailing coordinates.
				Else If(test.strLocatorDic.ContainsKey(objUserArgsList.Item(intIndex))) Then
					
					'Do a simple translate and append slashes if needed
					objUserArgsList.Item(intIndex) = FnReplaceImproperSpaces(FnMakeLeadingSlashes(test.strLocatorDic.Item(objUserArgsList.Item(intIndex))))
					
					'Try to get the position of the pattern matching "@[" in case this is a coordinate-based object
					intTemp = InStrRev(objUserArgsList.Item(intIndex), "@[")
					
					'If the control string ends with @[##:##]
					If intTemp > 0 AndAlso FnIsCoordinate(Mid(objUserArgsList.Item(intIndex), intTemp+1), x, y) Then
						
						'Strip the suffix of this object  
						objUserArgsList.Item(intIndex) = Left(objUserArgsList.Item(intIndex), intTemp-1)
						
						'Get the xpath of the object at the specified coordinate
						objUserArgsList.Item(intIndex) = FnGetXpathFromCoordinateWithinParent(objUserArgsList.Item(intIndex), New SilkTest.Ntf.Point(x,y))
							
					End If	
					
					'If the user isn't invoking the SetParent command, then we may need to tack on the current xpath parent
					If String.Compare(strCmd, "setparent", True) <> 0 Then
					
						'If the ctrlName doesn't contain the test.strParentXpath value, then tack on the test.strParentXpath value
						If Len(test.strParentXpath) > 0 AndAlso objUserArgsList.Item(intIndex).ToLower.Contains(test.strParentXpath.ToLower) = False Then
							objUserArgsList.Item(intIndex) = test.strParentXpath & objUserArgsList.Item(intIndex)
						
						'Else if the ctrlName is equal to the test.strParentXpath value and if the command is "Close", then we want to wipe out the test.strParentXpath value entirely
						Else If Len(test.strParentXpath) > 0  AndAlso String.Equals(objUserArgsList.Item(intIndex).ToLower, test.strParentXpath.ToLower) AndAlso String.Compare(strCmd, "close", True) = 0 Then
							test.strParentXpath = ""
						End If
						
					End If
					
				'Otherwise, if this parameter starts with '$' AND [ends in empty brackets '[]' OR [is not present in test.strUserVarDic AND does not contain left/right brackets]] THEN 
				'we know it's a not-yet-existant variable name to be populated.  Create/update it.
				Else If(String.Equals(Left(objUserArgsList.Item(intIndex), 1), "$") AndAlso (String.Equals(Right(objUserArgsList.Item(intIndex), 2), "[]") OrElse _
					   (test.strUserVarDic.ContainsKey(Mid(objUserArgsList.Item(intIndex), 2))=False AndAlso objUserArgsList.Item(intIndex).Contains("[")=False AndAlso objUserArgsList.Item(intIndex).Contains("]")=False))) Then
								
					'Strip the leading dollar sign because it's no longer necessary
					objUserArgsList.Item(intIndex) = Mid(objUserArgsList.Item(intIndex), 2)
					
					'If the param is less than 3 characters in length or does not end in empty brackets, just create/overwrite a variable without declaring type
					If(Len(objUserArgsList.Item(intIndex)) < 3 OrElse Not String.Equals(Right(objUserArgsList.Item(intIndex), 2), "[]")) Then
						If(Not test.strUserVarDic.ContainsKey(objUserArgsList.Item(intIndex))) Then
							test.strUserVarDic.Add(objUserArgsList.Item(intIndex), "placeHolder")
						Else
							test.strUserVarDic.Item(objUserArgsList.Item(intIndex)) = "placeHolder"
						End If
						
					'Else if the param is 5 characters or greater and ends with '[][]', then drop and last 4 characters and create/overwrite a 2D associative array
					Else If(Len(objUserArgsList.Item(intIndex))>=5 AndAlso String.Equals(Right(objUserArgsList.Item(intIndex), 4), "[][]")) Then
						objUserArgsList.Item(intIndex) = Left(objUserArgsList.Item(intIndex), Len(objUserArgsList.Item(intIndex))-4)
						If(Not test.strUserVarDic.ContainsKey(objUserArgsList.Item(intIndex))) Then
							test.strUserVarDic.Add(objUserArgsList.Item(intIndex), New AssociativeArray2D)
						Else
							test.strUserVarDic.Item(objUserArgsList.Item(intIndex)) = New AssociativeArray2D
						End If
						
					'Else this param must end in '[]', so drop the last 2 characters and create/overwrite a dictionary object
					Else
						objUserArgsList.Item(intIndex) = Left(objUserArgsList.Item(intIndex), Len(objUserArgsList.Item(intIndex))-2)
						If(Not test.strUserVarDic.ContainsKey(objUserArgsList.Item(intIndex))) Then
							test.strUserVarDic.Add(objUserArgsList.Item(intIndex), New Dictionary(Of String, String)(System.StringComparer.OrdinalIgnoreCase))
						Else
							test.strUserVarDic.Item(objUserArgsList.Item(intIndex)) = New Dictionary(Of String, String)(System.StringComparer.OrdinalIgnoreCase)
						End If
					End If
					
				'Otherwise, IF [this param starts and ends with double-quotes] OR [begins with '$'] OR [starts with '#' AND [is in dicConstants OR is a fixed constant]] THEN send it to the substitutor
				Else If((String.Equals(Left(objUserArgsList.Item(intIndex), 1), """") AndAlso String.Equals(Right(objUserArgsList.Item(intIndex), 1), """")) OrElse _ 
						(String.Equals(Left(objUserArgsList.Item(intIndex), 1), "$")) OrElse _
						(String.Equals(Left(objUserArgsList.Item(intIndex), 1), "#") AndAlso (dicConstants.ContainsKey(Mid(objUserArgsList.Item(intIndex), 2)) OrElse String.Equals(Mid(objUserArgsList.Item(intIndex), 2), dicConstants.Item("CurrentDateTimeVarName"))))) Then
						
					'Assuming the user isn't trying to overwrite an existing variable, send the param to the substitutor
					If(Not( _
						(intIndex=2 And String.Compare(strCmd, "Retain", True)=0) OrElse _
						(intIndex=3 And String.Compare(strCmd, "Math", True)=0) OrElse _
						(intIndex=3 And String.Compare(strCmd, "SQL", True)=0) OrElse _
						(intIndex=0 And String.Compare(strCmd, "Declare", True)=0) _
					))
					
						'Send it through the substitutor
						objUserArgsList.Item(intIndex) = FnVarSub(objUserArgsList.Item(intIndex), test, dicConstants)
						
						'If the user is manually sending an xpath, we need to swap out invalid spaces at this time
						If FnIsXpath(objUserArgsList.Item(intIndex)) Then 
							objUserArgsList.Item(intIndex) = test.strParentXpath & FnReplaceImproperSpaces(objUserArgsList.Item(intIndex))
						End If
					
					'Else, the user is actually overwriting an existing parameter.  Just strip the dollar sign and move on.
					Else
						objUserArgsList.Item(intIndex) = Mid(objUserArgsList.Item(intIndex), 2)
					End If						
					
				'Otherwise, if the param contains the '@' symbol followed by the ':' character and stripping everything from the '@' symbol onward matches a key in test.strLocatorDic, then strip/retain the suffix and translate the param
				'NOTE: This logic does not currently account for the instance where a row or column label contains an '@' symbol.  The interpreter will always say that it couldn't find a definition for your table object.
				'NOTE: It also fails to account for the instance where row or column labels contain the ':' character.  This will eventually need to be updated.
				'NOTE: This will only work for HTML tables at the moment.  Window.Table doesn't have a .GetCell method, so I don't know how to retrieve the specified cell's control name.  I need a Window.Table example to identify against.
				Else If InStrRev(objUserArgsList.Item(intIndex), "@") > 0 AndAlso InStrRev(objUserArgsList.Item(intIndex), "@") < InStrRev(objUserArgsList.Item(intIndex), ":") AndAlso _
						test.strLocatorDic.ContainsKey(Left(objUserArgsList.Item(intIndex), InStrRev(objUserArgsList.Item(intIndex), "@")-1)) Then
						
					'Split the whole parameter by the '@' symbol to more easily consume it
					strTempArr = Split(objUserArgsList.Item(intIndex), "@")
					
					'Now split the last element in the above array by the ':" character
					strTempArr = Split(strTempArr(strTempArr.Length-1), ":")
					
					'Strip the suffix
					objUserArgsList.Item(intIndex) = Left(objUserArgsList.Item(intIndex), InStrRev(objUserArgsList.Item(intIndex), "@")-1)
					
					'Translate the table/grid
					objUserArgsList.Item(intIndex) = FnReplaceImproperSpaces(FnMakeLeadingSlashes(test.strLocatorDic.Item(objUserArgsList.Item(intIndex))))
					
					'Now get the control name of the cell itself
					objUserArgsList.Item(intIndex) = FnGetCellXpath(test.strParentXpath & objUserArgsList.Item(intIndex), strTempArr(0), strTempArr(1), dicConstants, FnInterpret)
									
				'If there is no translation for this unquoted string, we have a problem.
				Else
					FnInterpret = "The interpreter could find no corresponding definition for parameter: '" & objUserArgsList.Item(intIndex) & "'.  " & _
								  "Please ensure you've loaded the needed objects via the LoadObjects command.  " & _
								  "If this is a loose string not meant to be correlated in an existing Excel doc or data type, you need to enclose the parameter in double-quotes.  " & _
								  "If it appears that your parameter has been truncated, you may have used a comma outside of double-quotes by accident."
				End If
				
				'Troubleshooting
				'Console.WriteLine("The interpreted parameter is: " & objUserArgsList.Item(intIndex))
				
			Next
		End If
		
		'If this instruction has a trailing hyphen and isn't the Wait command, then cycle through the parameters and look for anything beginning with "//".
		'If you find a match, ensure the object is available before proceeding.  Otherwise, Xscript will appear to hang for ~10 seconds.
		If String.IsNullOrEmpty(FnInterpret) AndAlso blnTrailingHyphen AndAlso String.Compare(strCmd, "Wait", True)<>0 Then
			For Each oneParam As String In objUserArgsList
				If String.Equals("//", Left(oneParam, 2)) AndAlso FnExists(oneParam)=False Then
					blnOccasionalObjectNotPresent = True
					Exit For
				End If
			Next oneParam
		End If
		
		'Assuming an error hasn't already occurred and this step isn't trying to interact with an occasional object
		If(String.IsNullOrEmpty(FnInterpret) AndAlso blnOccasionalObjectNotPresent=False) Then
													
			'Tack on a bunch of empty items to our objUserArgsList.  This is so we don't have to check for list size before calling FnVerify, FnClick, FnType, etc.  We can just send it all.
			'We would simply send each of these functions the list/array of commands, but that will make for sloppy function definitions which are hard to utilize in future applications.
			For i As Integer = 1 To intNumNothings
				objUserArgsList.Add(Nothing)
			Next i
						
			'Figure out what the command is and send it off to its appropriate function.  If the corresponding function returns a string value, then we know an error has occurred.
			Select Case strCmd.ToUpper
				
				'=========== Browser =============
				Case "DISMISSDOWNLOAD"
					FnInterpret = FnDismissDownload
				Case "SAVEDOWNLOAD"
					FnInterpret = FnSaveDownload(objUserArgsList.Item(0))
				Case "NAVIGATE"
					FnInterpret = FnNavigate(objUserArgsList.Item(0))
					
				'======== Xscript Core ==========
				Case "SETPARENT"
					If objUserArgsList.Count-intNumNothings <= 1 OrElse objUserArgsList.Item(1) = False Then 
						test.strParentXpath = ""
					End If
					test.strParentXpath = test.strParentXpath & objUserArgsList.Item(0)
				Case "VERIFY"
					FnInterpret = FnVerify(objUserArgsList.Item(0), blnInvertTest, dicConstants, objUserArgsList.Item(1), objUserArgsList.Item(2), If(objUserArgsList.Item(3) Is Nothing, False, objUserArgsList.Item(3)))
				Case "RETAIN"
					FnInterpret = FnRetain(objUserArgsList.Item(0), objUserArgsList.Item(1), test.strUserVarDic.Item(objUserArgsList.Item(2)), dicConstants, objUserArgsList.Item(3))
				Case "WAIT"
					FnInterpret = FnWait(objUserArgsList.Item(0), objUserArgsList.Item(1), objUserArgsList.Item(2), objUserArgsList.Item(3), objUserArgsList.Item(4), dicConstants)
				Case "DECLARE"
					If test.strUserVarDic.ContainsKey(objUserArgsList.Item(0)) Then test.strUserVarDic.Item(objUserArgsList.Item(0))=objUserArgsList.Item(1) Else test.strUserVarDic.Add(objUserArgsList.Item(0), objUserArgsList.Item(1))
				
				'======= Xscript Compare ========
				Case "EQUALS"
					FnInterpret = FnEquals(objUserArgsList.Item(0), objUserArgsList.Item(1), objUserArgsList.Item(2))
				Case "NOTEQUALS"
					FnInterpret = FnNotEquals(objUserArgsList.Item(0), objUserArgsList.Item(1), objUserArgsList.Item(2))
				Case "MATH"
					FnInterpret = FnMath(objUserArgsList.Item(0), objUserArgsList.Item(1), objUserArgsList.Item(2), test.strUserVarDic.Item(objUserArgsList.Item(3)))
				Case "EXCELCOMPARE"
					FnInterpret = FnCompareExcel(objUserArgsList.Item(0), objUserArgsList.Item(1), objUserArgsList.Item(2))
				Case "VERIFYIMAGE"
					SubVerifyImage(dicConstants.Item("BitmapChecksPath") & "\" & objUserArgsList.Item(0), If(objUserArgsList.Item(1) Is Nothing, dicConstants.Item("DefaultBitmapCheckTolerancePercentage"), objUserArgsList.Item(1)), _
													 If(objUserArgsList.Item(2) Is Nothing, "red", objUserArgsList.Item(2)), test, dtCurrentDateTime, dicConstants)
					
					
				'======= Xscript System =========
				Case "INSERTSTEPS"
					FnInterpret = FnInsertSteps(test, objUserArgsList(0), dicConstants)
				Case "RUN"
					strTemp = Path.GetFileName(objUserArgsList(0)).ToUpper
					FnInterpret = FnRun(objUserArgsList.Item(0), If(strActiveExeFileNameList.Contains(strTemp) And objUserArgsList.Item(1), True, False), objUserArgsList.Item(2))
					If Not strActiveExeFileNameList.Contains(strTemp) Then strActiveExeFileNameList.Add(strTemp)
				Case "LOADOBJECTS"
					FnInterpret = FnLoadObjects(objUserArgsList, test, dicConstants)
				Case "STARTLOOP"
					FnInterpret = FnStartLoop(test, objUserArgsList.Item(0), dicConstants, objUserArgsList.Item(1), objUserArgsList.Item(2), objUserArgsList.Item(3), objUserArgsList.Item(4))
				Case "ENDLOOP"
					FnInterpret = FnEndLoop(test, objUserArgsList.Item(0))		
				Case "OUTPUTTEXT"
					test.strOutputTextList.Add(objUserArgsList.Item(0))
				Case "SQL"
					If objUserArgsList.Item(3) Is Nothing
						FnInterpret = FnSQL(objUserArgsList.Item(0), objUserArgsList.Item(1), objUserArgsList.Item(2))
					Else
						FnInterpret = FnSQL(objUserArgsList.Item(0), objUserArgsList.Item(1), objUserArgsList.Item(2), test.strUserVarDic.Item(objUserArgsList.Item(3)), objUserArgsList.Item(4))
					End If
					
				'========== Invocation ==========
				Case "CLICK", "DOUBLECLICK", "MOUSEMOVE", "PRESSKEYS", "RELEASEKEYS", "PRESSMOUSE", "RELEASEMOUSE", "SETTEXT", "TYPEKEYS", "SELECT", "CLOSE", "SETACTIVE", "SETFOCUS", "TEXTCLICK"
					FnInterpret = FnTestObjectInvoke(strCmd, dicConstants, objUserArgsList.GetRange(0, objUserArgsList.Count-intNumNothings))
				Case Else
					If Not FnCallCustomMethod(strCmd, objUserArgsList.GetRange(0, objUserArgsList.Count-intNumNothings)) Then
						FnInterpret = "Xscript does not recognize the '" & strCmd & "' command.  If this is a custom function or subroutine built outside of Xscript, " & _
									  "please ensure you have added the corresponding .NET script into the 'Xscript Caller' script.  Then, ensure that you have added a " & _
									  "direct reference to its unique class/module name within the 'Xscript Caller' script where it reads 'ADD THE NAME OF YOUR MODULE HERE'."
					End If
					
			End Select
			
		End If
		
	End Function
	

'==============================================================================================================
' FnStartLoop
'
' 		Function is called when the 'StartLoop' command is executed in Xscript.  It is responsible for creating
'		new Test objects for each column of data in the corresponding data worksheet.  Only one Test object is
'		created upon calling this function; the function will end up being called over and over because Xscript
'		sets the current command back to this 'StartLoop' command upon reaching the corresponding 'EndLoop'.
'
'		Though a bit counterintuitive, FnStartLoop is responsible for finally exiting the loop block when the
'		time comes.  When all loops are exhausted, it sets the current instruction to its corresponding 
'		EndLoop command.  This 'tricks' Xscript into thinking that it just executed the EndLoop command and 
'		will then move on to the subsequent command, thereby removing us from the loop block.
'
'		It is possible to have correlated loops by using all the optional parameters.  If the value passed in
'		for strVarValue1 and the value represented by the variable named strVarName2 are equal, then this function
'		will create a Test object for that particular column of data.  If those two values are not equal, then
'		the corresponding column will be ignored.
'
'		@test:							The current test
'		@strSheetName:					The name of the sheet that contains the user's data
'		@dicConstants:					The dictionary of Xscript constants derived from Constants.xlsx
'		@objBlnReportEachLoopResult:	A boolean cast as an object to represent whether each loop should create a new row in the results
'		@objBlnAbortOnFail:				When a loop iteration finds a failure, should Xscript keep chugging or reset back to step #1
'		@strVarValue1:					If user wants a correlated loop, this is the first of two values to correlate
'		@strVarName2:					If user wants a correlated loop, this string represents the name of the second variable to correlate.
'==============================================================================================================	

	Private Function FnStartLoop(ByRef test As Test, _
								 ByVal strSheetName As String, _
								 ByRef dicConstants As Dictionary(Of String, String), _
								 Optional ByVal objBlnReportEachLoopResult As Object=Nothing, _
								 Optional ByVal objBlnAbortOnFail As Object=Nothing, _ 
								 Optional ByVal strVarValue1 As String=Nothing, _
								 Optional ByVal strVarName2 As String=Nothing) As String
		
		'Variables
		Dim testChild As Test
		Dim intEndLoopIndex As Integer
		Dim blnHasUnfinishedChildren As Boolean
		
		'Default return
		FnStartLoop=Nothing
		
		'Immediately lowercase the sheetName.  It's a royal pain otherwise.
		strSheetName = strSheetName.ToLower
		
		'Find the corresponding EndLoop instruction
		intEndLoopIndex = FnLowerCaseList(test.strStepsList).IndexOf("endloop(""" & strSheetName.ToLower & """)")
		
		'Ensure there is a corresponding EndLoop command
		If intEndLoopIndex > 0 And test.intCurrentStepIndex+1 < intEndLoopIndex Then
		
			'Ensure the second sheet actually exists
			If FnSheetExistsInCache(test.strExcelFilePath, strSheetName) Then
				
				'If we haven't already loaded this particular datasheet, load it
				If Not test.dataWorksheetDic.ContainsKey(strSheetName) Then
					test.dataWorksheetDic.Add(strSheetName.ToLower, FnCreateComplexTDAA(test.strExcelFilePath, strSheetName, False, True, True))
				End If
				
				'Loop across every column 
				For Each strColumnKey As String In test.dataWorksheetDic(strSheetname).Keys
					
					'If this column is active AND we haven't already created a child Test object for it AND [user does not want a coordinated loop OR he does and this column is properly coordinated]
					If Not String.IsNullOrEmpty(test.dataWorksheetDic(strSheetname).Item(strColumnKey, "Active")) AndAlso test.testChildDic.ContainsKey(strColumnKey)=False AndAlso _
					   ((strVarValue1 Is Nothing OrElse strVarName2 Is Nothing) OrElse String.Compare(strVarValue1, test.dataWorksheetDic(strSheetname).Item(strColumnKey, strVarName2), True)=0) Then
					
						'Create the child test object
						testChild = New Test(test.strTestName & " - " & strColumnKey, test.strExcelFilePath)
						
						'Give it a reference to its parent test object
						testChild.testParent = test
						
						'Tell the child where it was born within the parent
						testChild.intParentStartLoopIndex = test.intCurrentStepIndex
						
						'Tell the child its true name
						testChild.strLoopColumnHeader = strColumnKey
						
						'Set behavior booleans based on StartLoop args
						testChild.blnReportResults = If(objBlnReportEachLoopResult Is Nothing, True, objBlnReportEachLoopResult)
						testChild.blnAbortOnFail = If(objBlnAbortOnFail Is Nothing, True, objBlnAbortOnFail)
						
						'Create a short list of steps from the parent's list of steps
						testChild.strStepsList.AddRange(test.strStepsList.GetRange(test.intCurrentStepIndex+1, intEndLoopIndex-test.intCurrentStepIndex-1))
						
						'Set strUserVarDic by looping through every variable in the datasheet at the appropriate column
						For Each strVarName As String In test.dataWorksheetDic(strSheetname).Keys(strColumnKey)
							
							'If this variable represents an existing constant, substitute it now
							If Len(test.dataWorksheetDic(strSheetname).Item(strColumnKey, strVarName))>0 AndAlso _
							   String.Equals(Left(test.dataWorksheetDic(strSheetname).Item(strColumnKey, strVarName), 1), "#") AndAlso _
							   dicConstants.ContainsKey(Mid(test.dataWorksheetDic(strSheetname).Item(strColumnKey, strVarName), 2)) Then
								test.dataWorksheetDic(strSheetname).Update(strColumnKey, strVarName, dicConstants.Item(Mid(test.dataWorksheetDic(strSheetname).Item(strColumnKey, strVarName), 2)))
							End If
							
							'Depending on whether the user constant is already taken, insert or overwrite it now
							If Not testChild.strUserVarDic.ContainsKey(strVarName) Then
								testChild.strUserVarDic.Add(strVarName, test.dataWorksheetDic(strSheetname).Item(strColumnKey, strVarName))
							Else
								testChild.strUserVarDic.Item(strVarName) = test.dataWorksheetDic(strSheetname).Item(strColumnKey, strVarName)
							End If
							
						Next strVarName
						
						'Add reference to this test child in the parent test						
						test.testChildDic.Add(strColumnKey, testChild)
						
						'Stop looping through the columns because we don't want to create more child tests until they're needed
						Exit For
						
					End If
					
				Next strColumnKey
				
				'Loop through all this test's children and figure out if any of them are unfinished.
				For Each kvpChildTest As KeyValuePair(Of String, Test) In test.testChildDic				
					If kvpChildTest.Value.blnTestDone = False Then
						blnHasUnfinishedChildren = True
						Exit For						
					End If
				Next kvpChildTest
				
				'If the test has no unfinished children, then we can set its current step to the corresponding EndLoop command.  Back in the 'Xscript' script, it will advance
				'the intCurrentStepIndex as usually.  By setting the step here, we guarantee that the 'Xscript' will cause us to finally escape the loop.  This is the only way out.
				If blnHasUnfinishedChildren=False Then
					test.intCurrentStepIndex = intEndLoopIndex
				End If
				
			End If
			
		Else If test.intCurrentStepIndex+1=intEndLoopIndex Then
			FnStartLoop = "The StartLoop command is immediately followed by its corresponding EndLoop command.  This is an invalid use of loops; you must have instruction in between."
		Else If intEndLoopIndex = -1
			FnStartLoop = "There is no corresponding 'EndLoop(""" & strSheetName & """)' instruction."
		End If
	
	End Function	

'==============================================================================================================
' FnEndLoop
'
' 		FnEndLoop is executed when Xscript encounters the 'EndLoop' command.  Its only job is to set the
'		current instruction back to the corresponding 'StartLoop' command.  
'
' 		@test: 			The currently executing test
' 		@strSheetName: 	The name of the datasheet containing all the loop variables
'==============================================================================================================	

	Private Function FnEndLoop(ByRef test As Test, ByVal strSheetName As String) As String
	
		'Declarations
		Dim intStartLoopIndex As Integer
	
		'Default return
		FnEndLoop=Nothing
	
		'Make sure you're at least on the 3rd step 
		If test.intCurrentStepIndex >= 2 Then
		
			'Immediately lowercase the sheetName.  It's a royal pain otherwise.
			strSheetName = strSheetName.ToLower
			
			'Find the corresponding StartLoop instruction
			intStartLoopIndex = FnLowerCaseList(test.strStepsList).IndexOf("startloop(""" & strSheetName.ToLower & """)")
		
				'Make sure there's a corresponding 'StartLoop' command at least two steps backward
				If intStartLoopIndex > -1 And intStartLoopIndex+1 < test.intCurrentStepIndex Then
					
					'Set the test's current step index such the next step is the corresponding startloop command
					test.intCurrentStepIndex = intStartLoopIndex-1
					
				Else If intStartLoopIndex > -1 Then
					FnEndLoop = "The EndLoop command must have a corresponding StartLoop command referencing the same worksheet."
				Else If intStartLoopIndex+1 = test.intCurrentStepIndex
					FnEndLoop = "The EndLoop command cannot immediately follow the StartLoop command.  Minimally, there must be one step in between."
				Else
					FnEndLoop = "The EndLoop command must come *after* the corresponding StartLoop command."
				End If
		Else
			FnEndLoop = "The EndLoop command cannot be the first or second command in a test.  Minimally, it must be preceeded by the corresponding StartLoop command and an instruction."
		End If
	
	End Function
	
		
'==============================================================================================================
' FnInsertSteps
'
' 		Executes upon reaching 'InsertSteps' command.  It loads the corresponding Excel sheet via FnLoadSimpleList,
'		then it sticks all the contained steps directly into the Test's list of steps.  Easy peasy.
'
' 		@test: 			The currently executing test
' 		@strFilePath: 	Full file path to the Excel file containing all steps to insert
'		@dicConstants:	A dictionary of variables that help tweak behavior in the Xscript script
'==============================================================================================================	

	Private Function FnInsertSteps(ByRef test As Test, ByVal strFilePath As String, ByRef dicConstants As Dictionary(Of String, String)) As String
	
		'Objects
		Dim strInsertionStepsList As New List(Of String)
		
		'Default return
		FnInsertSteps = ""
	
		'Get the list to insert
		strInsertionStepsList = FnLoadSimpleList(dicConstants.Item("LibraryLocation") & "\" & strFilePath, dicConstants.Item("InstructionsColumnHeader"), True, True)
							
		'Insert new steps into our current steps
		test.strStepsList.InsertRange(test.intCurrentStepIndex+1, strInsertionStepsList)
								
		'Retain how many steps have been added
		test.intNumStepsAdded = test.intNumStepsAdded + strInsertionStepsList.Count
								
	End Function
	
'==============================================================================================================
' FnLoadObjects
'
' 		Executed whenever the LoadObjects command is encountered by Xscript.  It is responsible for retrieving
'		the referenced Excel file of objects and loading them all into the Test object.  This function ensures
'		that the Default.xlsm objects workbook is always loaded.  This function is also responsible for
'		calling FnUpdateObjectReferences for both the referenced object workbook and Default.xlsm.  Note, however,
'		that FnUpdateObjectReferences is quickly aborted when there are no values in the 'New coded name' columns.
'
' 		@objUserArgsList: 	All of the arguments passed in by the user (post interpretation)
' 		@test: 				The currently executing test
'		@dicConstants:		A dictionary of variables that help tweak behavior in the Xscript script
'==============================================================================================================
	
	Private Function FnLoadObjects(ByRef objUserArgsList As List(Of Object), ByRef test As Test, ByRef dicConstants As Dictionary(Of String, String)) As String
	
		'Default return
		FnLoadObjects=Nothing
	
		'If we haven't already loaded any items into the locator dictionary
		If test.strLocatorDic.Count=0 Then
			
			'Send the default objects workbook into the FnUpdateObjectReferences routine in case there are updates to be made
			FnLoadObjects = FnUpdateObjectReferences(dicConstants.Item("DefaultObjectsWorkbook"), test, dicConstants)
			
			'Now, load the default objects workbook
			test.strLocatorDic = FnLoadSimpleDictionary(FnLoadObjects, dicConstants.Item("ObjectsFolder") & "\" & dicConstants.Item("DefaultObjectsWorkbook"), _
														dicConstants.Item("ObjectNameColumnHeader"), dicConstants.Item("ObjectValueColumnHeader"), True)
			
		End If
		
		'If we haven't already encountered an error
		If String.IsNullOrEmpty(FnLoadObjects) Then
			
			'Call FnUpdateObjectReferences.  This updates every active test, inactive test, and library file to have the new reference.  It also updates the given Objects sheet.  It updates test.strStepsList, too.
			FnLoadObjects = FnUpdateObjectReferences(objUserArgsList.Item(0), test, dicConstants)
			
			'If we haven't already encounter an error
			If String.IsNullOrEmpty(FnLoadObjects) Then
				
				'If this test's strLocatorDic is empty, set it to the values found in the supplied objects worksheet file (it will only be empty at this point if the default objects worksheet is empty)
				If test.strLocatorDic.Count = 0 Then
					test.strLocatorDic = FnLoadSimpleDictionary(FnLoadObjects, dicConstants.Item("ObjectsFolder") & "\" & objUserArgsList.Item(0), dicConstants.Item("ObjectNameColumnHeader"), _
																dicConstants.Item("ObjectValueColumnHeader"), True, If(objUserArgsList(1) Is Nothing, Nothing, objUserArgsList.GetRange(1, objUserArgsList.Count-intNumNothings-1).ToArray))
			
				'If this test's strLocatorDic has value, load the values within the given objects file and allow the new objects to overwrite the old objects											
				Else
					For Each kvp As KeyValuePair(Of String, String) In FnLoadSimpleDictionary(FnLoadObjects, dicConstants.Item("ObjectsFolder") & "\" & objUserArgsList.Item(0), dicConstants.Item("ObjectNameColumnHeader"), _
																	   dicConstants.Item("ObjectValueColumnHeader"), True, If(objUserArgsList(1) Is Nothing, Nothing, objUserArgsList.GetRange(1, objUserArgsList.Count-intNumNothings-1).ToArray))
						If Not test.strLocatorDic.ContainsKey(kvp.Key) Then
							test.strLocatorDic.Add(kvp.Key, kvp.Value)
						Else
							test.strLocatorDic.Item(kvp.Key) = kvp.Value
						End If
					Next kvp
				End If
				
			End If
			
		End If
					
	End Function
	
'==============================================================================================================
' FnIsCoordinate
'
' 		Determines whether a given string represents a coordinate.  Xscript defines a coordinate as an unquoted
'		string surrounded in brackets which contain two colon-delimited integers.  For example: [46:812].
'
'		If the input is determined to be a coordinate, this function will also populate the given 'x' and 'y'
'		variables.  If it's not a coordinate, then both variables are populated with -1.
'
' 		@strInput: 	The string input to be evaluated
' 		@x: 		ByRef variable to be populated with the x coordinate
' 		@y: 		ByRef variable to be populated with the y coordinate
'==============================================================================================================
	
	Private Function FnIsCoordinate(ByVal strInput As String, ByRef x As Integer, ByRef y As Integer) As Boolean
	
		'Default return
		FnIsCoordinate = False
	
		'If it starts and ends with brackets
		If(String.Equals("[", Left(strInput, 1)) AndAlso String.Equals("]", Right(strInput, 1))) Then
		
			'Strip off the brackets and trim the remainder
			strInput = Mid(strInput, 2, Len(strInput)-2).Trim
			
			'Split the remainder by the colon
			strTempArr = Split(strInput, ":")
			
			'If there are two elements in the array and they're both integers, set x and y
			If (strTempArr.length = 2 AndAlso Integer.TryParse(strTempArr(0), x) AndAlso Integer.TryParse(strTempArr(1), y)) Then
				FnIsCoordinate = True
			Else
				x=-1
				y=-1
			End If
			
		End If
				
	End Function
	

'==============================================================================================================
' FnIsXpath
'
'		Returns TRUE if the supplied string represents an xpath.  It must start with double slashes and contain
'		a left bracket followed by a right bracket.  Really I should be doing this with regex, but I'm not 
'		readily versed in regex and this will do just fine instead for now.
'
' 		@strInput: 	The string input to be evaluated
'==============================================================================================================
	
	Public Function FnIsXpath(ByVal strInput As String) As Boolean
		
		FnIsXpath = False
		
		strInput = strInput.Trim
		
		If Len(strInput) >= 6 AndAlso String.Equals("//", Left(strInput, 2)) AndAlso InStr(strInput, "[") > 0 AndAlso _
		   InStrRev(strInput, "]") > 0 AndAlso InStr(strInput, "[") < InStrRev(strInput, "]") Then
			FnIsXpath = True
		End If
						
	End Function

'==============================================================================================================
' FnGetActiveExeFileNameList
'
' 		Simple 'Get' function which returns the list of currently running executables that Xscript is responsible
'		for kicking off.  This list is appended every time the 'Run' command is sent to Xscript.  The reason
'		we need a 'Get' function is because the main 'Xscript' script needs to retrieve this list and the Interpreter
'		is a module and not a class, so its properties are automatically private.  The Xscript script needs this
'		list in case it catches and exception and needs to kill all running applications to ensure stability.
'==============================================================================================================
	
	Public Function FnGetActiveExeFileNameList As List(Of String)
		Return strActiveExeFileNameList
	End Function

'==============================================================================================================
' SubClearActiveExeFileNameList
'
' 		Simple 'Set' routine that clears the list of currently running executables kicked off by Xscript.  This
'		is needed so that the Xscript script can reset the list upon catching an exception and killing all the
'		applications it was responsible for kicking off.  The reason we need a 'Set' function is because the 
'		Interpreter is a module and not a function, therefore we could not delcare strActiveExeFileNameList as
'		public.  I think.  Eh just go with it.
'==============================================================================================================
	
	Public Sub SubClearActiveExeFileNameList
		strActiveExeFileNameList.Clear
	End Sub

'==============================================================================================================
' FnTranslateMouseButton
'
' 		Many native Silk functions utilize the MouseButton enumeration to send left, right or middle mouse clicks.
'		Because I am invoking those functions via Reflection, I need to ensure I feed them the proper enumeration.
'		To do this, I need to capture unquoted strings that look like references to the MouseButton enumeration
'		and return the actual enumeration object instead.
'
' 		@strInput: 		User-supplied string input to evaluate
' 		@strErrorMsg: 	If an error occurs, this is the ByRef variable to populate with the error message
'==============================================================================================================
		
	Private Function FnTranslateMouseButton(ByVal strInput As String, ByRef strErrorMsg As String) As Object	
		FnTranslateMouseButton = ""
		Select Case strInput.ToLower
			Case "mousebutton.left"
				FnTranslateMouseButton = MouseButton.Left
			Case "mousebutton.right"
				FnTranslateMouseButton = MouseButton.Right
			Case "mousebutton.middle"
				FnTranslateMouseButton = MouseButton.Middle
			Case Else
				strErrorMsg = "The parameter '" & strInput & "' is not recognized as a valid MouseButton enumeration.  " & _
							  "Acceptable values are 'Left', 'Right', and 'Middle'.  " & _
							  "If this was Not meant To be a MouseButton enumeration, place Double-quotes around the parameter."
		End Select				
	End Function
		
		
'==============================================================================================================
' FnTranslateClickType
'
' 		Some native Silk functions utilize the ClickType enumeration to send different types of clicks to an object.
'		Because I am invoking those functions via Reflection, I need to ensure I feed them the proper enumeration.
'		To do this, I need to capture unquoted strings that look like references to the ClickType enumeration
'		and return the actual enumeration object instead.
'
' 		@strInput: 		User-supplied string input to evaluate
' 		@strErrorMsg: 	If an error occurs, this is the ByRef variable to populate with the error message
'==============================================================================================================
		
	Private Function FnTranslateClickType(ByVal strInput As String, ByRef strErrorMsg As String) As Object
		FnTranslateClickType = ""
		Select Case strInput.ToLower						
			Case "clicktype.left"
				FnTranslateClickType = ClickType.Left
			Case "clicktype.right"
				FnTranslateClickType = ClickType.Right
			Case "clicktype.middle"
				FnTranslateClickType = ClickType.Middle
			Case "clicktype.leftdouble"
				FnTranslateClickType = ClickType.LeftDouble
			Case "clicktype.press"
				FnTranslateClickType = ClickType.Press
			Case "clicktype.release"
				FnTranslateClickType = ClickType.Release
			Case Else
				strErrorMsg = "The parameter '" & strInput & "' is not recognized as a valid ClickType enumeration.  " & _
							  "Acceptable values are 'Left', 'Right', 'Middle', 'LeftDouble', 'Press', 'and Release'.  " & _
							  "If this was not meant to be a ClickType enumeration, place double-quotes around the parameter."
		End Select
	End Function
		
		
''==============================================================================================================
' FnTranslateModifierKeys
'
' 		Some native Silk functions utilize the ModifierKeys enumeration to send different keypresses during another action.
'		Because I am invoking those functions via Reflection, I need to ensure I feed them the proper enumeration.
'		To do this, I need to capture unquoted strings that look like references to the ModifierKeys enumeration
'		and return the actual enumeration object instead.
'
' 		@strInput: 		User-supplied string input to evaluate
' 		@strErrorMsg: 	If an error occurs, this is the ByRef variable to populate with the error message
'==============================================================================================================
		
	Private Function FnTranslateModifierKeys(ByVal strInput As String, ByRef strErrorMsg As String) As Object
		FnTranslateModifierKeys = ""
		Select Case strInput.ToLower						
			Case "modifierkeys.none"
				FnTranslateModifierKeys = ModifierKeys.None
			Case "modifierkeys.shift"
				FnTranslateModifierKeys = ModifierKeys.Shift
			Case "modifierkeys.alt"
				FnTranslateModifierKeys = ModifierKeys.Alt
			Case "modifierkeys.control"
				FnTranslateModifierKeys = ModifierKeys.Control
			Case "modifierkeys.controlalt"
				FnTranslateModifierKeys = ModifierKeys.ControlAlt
			Case "modifierkeys.shiftcontrol"
				FnTranslateModifierKeys = ModifierKeys.ShiftControl
			Case "modifierkeys.shiftalt"
				FnTranslateModifierKeys = ModifierKeys.ShiftAlt
			Case "modifierkeys.shiftcontrolalt"
				FnTranslateModifierKeys = ModifierKeys.ShiftControlAlt
			Case Else
				strErrorMsg = "The parameter '" & strInput & "' is not recognized as a valid ModifierKey enumeration.  " & _
							  "Acceptable values are 'ModifierKeys.None', 'ModifierKeys.Shift', 'ModifierKeys.Alt', 'ModifierKeys.Control', " & _
							  "'ModifierKeys.ControlAlt', 'ModifierKeys.ShiftControl', 'ModifierKeys.ShiftAlt', and 'ModifierKeys.ShiftControlAlt'.  " & _
							  "If this was not meant to be a ModifierKey enumeration, place double-quotes around the parameter."
		End Select
	End Function

		
'==============================================================================================================
' FnMakeLeadingSlashes
'
' 		When the interpreter matches an object reference in the object map, it calls this function to ensure
'		the corresponding xpath starts with two forward slashes.  This is needed in order for the SetParent
'		command to function properly because you can't append a parent xpath onto a child if the child isn't
'		prefixed with at least one slash.  The double forward slash says: "The object on the left contains the
'		object on the right, but it may not be an immediate relationship; there may be many generations in between."
'		If a single slash were used, that would indicate that it's a direct parent-child relationship with no 
'		generations in between - this would greatly reduce functionality of SetParent.
'
' 		@strInput: 	The xpath to be evaluated.
'==============================================================================================================
	
	Public Function FnMakeLeadingSlashes(ByVal strInput As String) As String
	
		FnMakeLeadingSlashes = strInput
		
		Do While String.Equals(Left(FnMakeLeadingSlashes, 2), "//") = False
			FnMakeLeadingSlashes = "/" & FnMakeLeadingSlashes
		Loop
		
	End Function						
		
'==============================================================================================================
' FnVarSub
'
' 		Substitutes all user variables and all constants found in a single string for their corresponding value.
'
'		User variables begin With '$' and Constants begin with '#'.  The function examines the given string from
'		right to left and locates the '$' or '#' symbols.  If the text following these symbols matches a value 
'       in dicConstants or the test.strUserVarDic collections, the variable name is substituted by its corresponding 
'		value.  This function properly handles the use of 1D and 2D array references and their index values, too.
'
'		This function is also responsible for translating the "fixed constant" notated by #DT.   This is not a 
'		constant that you'll find the Constants.xlsx dictionary; it is a constant the must be generated in real time.
'		It represents the current dateTime and is often used as a quasi random string.  It is formatted MMMddyyHHmmss
'		so that it begins with letters and ends with numbers.
'
' 		@strOneParam:	The parameter to evaluate for constants and userVars
' 		@test:			Currently executing Test object
' 		@dicConstants:	The dictionary of constants that may need to be substituted in the given string
'
'		Example #1
'			strMyParameter = "Welcome to the ODC, #ValidUserLoginName.  Today is $currentDate."
'			strMyParameter = FnVarSub(strMyParameter, dicConstants)
'			Console.Writeline(strMyParameter)
'				--> "Welcome to the ODC, Corveltest.  Today is April 17th, 2014."
'
'		Example #2
'			strMyParameter = "SELECT * FROM documents WHERE dcn = '$myDocTableRow[dcnColumn]'"
'			strMyParameter = FnVarSub(strMyParameter, dicConstants)
'			Console.Writeline(strMyParameter)
'				--> "SELECT * FROM documents WHERE dcn = '19832408'"
'
'		Example #3
'			strMyParameter = "Allentown has $cities[allentown][population] people living in it."
'			strMyParameter = FnVarSub(strMyParameter, dicConstants)
'			Console.Writeline(strMyParameter)
'				--> "Allentown has 350,000 people living in it."
'
'		Example #4
'			strMyParameter = "The #companyName company has $companies[#companyName][employeeCount] workers."
'			strMyParameter = FnVarSub(strMyParameter, dicConstants)
'			Console.Writeline(strMyParameter)
'				--> "The CorVel company has 3000 workers."
'==============================================================================================================
		
	Private Function FnVarSub(ByVal strOneParam As String, ByRef test As Test, ByRef dicConstants As Dictionary(Of String, String)) As String
		
		'Declare variables
		Dim intLastPossibleCandChar As Integer
		Dim i As Integer
		Dim intTemp As Integer
		Dim strCandidate As String
		Dim strTemp As String
		Dim strTempArr() As String
		
		'Set values
		FnVarSub = strOneParam
		intLastPossibleCandChar = -1
					
		'Remove leading and trailing double-quotes if necessary
		If(String.Equals(strOneParam(0), """") And String.Equals(Right(strOneParam,1), """"))
			FnVarSub = Mid(FnVarSub, 2, Len(FnVarSub)-2)
		End If
	
		'Start at the 2nd-last character in the string since a trailing # or $ char means nothing
		i = Len(FnVarSub)-2

		'Loop through the param char-by-char backwards
		Do While i >= 0
			
			'If we are on a dollar sign or pound sign
			If(FnVarSub.Chars(i) = "#" Or FnVarSub.Chars(i) = "$") Then
								
				'Figure out the index of the character that could be the last character in a variable name
				If(intLastPossibleCandChar = -1) Then intLastPossibleCandChar = Len(FnVarSub)-1
					
				'Loop backward from intLastPossibleCandChar
				Do While intLastPossibleCandChar > i
															
					'Create a candidate string for what the variable MIGHT be named
					strCandidate = Mid(FnVarSub, i+2, intLastPossibleCandChar-i)
					
					'If we started with a dollar sign
					If(FnVarSub.Chars(i) = "$") Then
												
						'If the candidate exists in the test.strUserVarDic dictionary
						If(test.strUserVarDic.ContainsKey(strCandidate))
														
							'If there are at least two characters available to the right of the last character 
							If(Len(FnVarSub) >= intLastPossibleCandChar + 3) Then
															
								'If the character to the right is a '[' bracket
								If(FnVarSub.Chars(intLastPossibleCandChar+1) = "[") Then
									
									'Get the index of the first ']' bracket to the right of the var name
									intTemp = Instr(Mid(FnVarSub, intLastPossibleCandChar+3), "]")
									
									'If the index is greater than zero, it means there's a right bracket
									If(intTemp > 0) Then
									
										'Redefine the strCandidate so it includes everything up to the aforementioned ']' bracket
										strCandidate = Mid(FnVarSub, i+2, intLastPossibleCandChar-i+intTemp+1)
										
										'If there are at least two characters available to the right of the last character 
										If(Len(FnVarSub) >= i + Len(strCandidate) + 3) Then
										
											'If the character to the right is a '[' bracket
											If(FnVarSub.Chars(i+Len(strCandidate)+1) = "[") Then
											
												'Get the index of the first ']' bracket to the right of the last ']' bracket
												intTemp = Instr(Mid(FnVarSub, i+len(strCandidate)+3), "]")
												
												'If the index is greater than zero, it means there's yet another right bracket
												If(intTemp > 0) Then
																									
													'Redefine the strCandidate so it includes everything up to the aforementioned ']' bracket
													strCandidate = Mid(FnVarSub, i+2, Len(strCandidate)+intTemp+1)
												
												End If
												
											End If
											
										End If
										
									End If
														
								End If
							
							End If														
							
							'Break down the string so we can extract the values within any existing brackets.
							'After these lines, we'll have three possible arrays: {myVarName}, {myVarName,myKey1}, {myVarName,myKey1,myKey2}
							strTemp = strCandidate.Replace("[", " ")
							strTemp = strTemp.Replace("]", "")
							strTemp = strTemp.Replace("""", "")
							strTempArr = Split(strTemp, " ")
													
							'If the key is present in our user's variable list, translate it
							If(test.strUserVarDic.ContainsKey(strTempArr(0))) Then
															
								'If there is only 1 element in the array, it's a simple string
								If(strTempArr.Length = 1) Then
									strTemp = test.strUserVarDic.Item(strTempArr(0))

								'If we were given only a var name and an index, the user is referencing a dictionary
								Else If(strTempArr.Length = 2) Then
									strTemp = test.strUserVarDic.Item(strTempArr(0)).Item(strTempArr(1))
									
								'If we were given a var name and two indices, the user is referencing a 2D associative array
								Else If(strTempArr.Length = 3) Then
									strTemp = test.strUserVarDic.Item(strTempArr(0)).Item(strTempArr(1), strTempArr(2))
								End If
								
							End If
							
							'Reconstruct the entire string so we can continue parsing it
							FnVarSub = Left(FnVarSub, i) & strTemp & Mid(FnVarSub, i + len(strCandidate) + 2)
							
							'Decrement i by an additional amount because there's no point in checking if the char to the left is a $ or # since it has no room to have a name assigned
							i = i - 1
							
							'Update the location of the last possible candidate char so we don't mistake the substituted text as part of a variable name
							intLastPossibleCandChar = i-1
							
							'Leave the current loop because we found a matching candidate and no longer need to look
							Exit Do
							
						End If
					
					'Else, the leading char must be a # symbol
					Else
					
						'If the candidate is present in the constants, sub it out
						If(dicConstants.ContainsKey(strCandidate)) 
							FnVarSub = Left(FnVarSub, i) & dicConstants.Item(strCandidate) & Mid(FnVarSub, i + len(strCandidate) + 2)
							
						'If the user is trying to get the current Date/Time, get the current date time and sub it out now
						Else If String.Equals(strCandidate, dicConstants.Item("CurrentDateTimeVarName")) Then							
							FnVarSub = Left(FnVarSub, i) & DateTime.Now.ToString("MMMddyyHHmmss") & Mid(FnVarSub, i + len(strCandidate) + 2)							
						End If
						
					End If
					
					'Decrement the length of the var name match string
					intLastPossibleCandChar = intLastPossibleCandChar - 1
					
				Loop				
				
			End If
			
			'Move backwards 1 character in the whole parameter
			i = i - 1
		
		Loop
		
	End Function
					
End Module






